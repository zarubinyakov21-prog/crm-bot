import imaplib, email, re, requests, os, json, logging, socket, sys
from email.header import decode_header
from bs4 import BeautifulSoup
from datetime import datetime
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# ===== НАСТРОЙКИ =====
IMAP_HOST = "imap.yandex.ru"
IMAP_PORT = 993
EMAIL_LOGIN = os.environ.get("EMAIL_LOGIN", "zakaz@olympickitchen.ru")
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
CRM_BASE_URL = "https://crm.private-crm.ru"
CRM_IDENTIFIER = "ok"
CRM_API_KEY = os.environ["CRM_API_KEY"]
CRM_PROJECT_ID = "1"
CRM_DELIVERY_TIME_ID = "20"  # fallback: 12:00-13:00

# /deliverytimes — API не поддерживает этот справочник, ID захардкожены вручную
_DELIVERY_TIMES = {
    "12:00-13:00": 20,
    "13:00-14:00": 21,
    "14:00-15:00": 22,
    "15:00-16:00": 23,
    "17:00-23:00": 5,
    "18:00-19:00": 24,
    "19:00-20:00": 25,
    "20:00-21:00": 26,
    "21:00-22:00": 27,
    "22:00-23:00": 28,
}
CRM_ORDER_SOURCE_ID = "1"
PROCESSED_IDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_processed_ids.json")
ERROR_LOG = os.path.expanduser("~/crm_bot/error.log")
SOCKET_TIMEOUT = 60
ADDR_MATCH_THRESHOLD = 0.4  # порог Jaccard для совпадения адресов
# =====================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(ERROR_LOG, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

socket.setdefaulttimeout(SOCKET_TIMEOUT)

_CRM_HEADERS = {"Identifier": CRM_IDENTIFIER, "Application-key": CRM_API_KEY}

# Стоп-слова при нормализации адреса (сокращения типов улиц, городов и т.п.)
_ADDR_STOP = {
    "г", "город", "ул", "улица", "пр", "проспект", "бул", "бульвар",
    "пер", "переулок", "пл", "площадь", "мкр", "микрорайон",
    "д", "дом", "кв", "квартира", "корп", "корпус", "стр", "строение",
    "ш", "шоссе", "пос", "поселок", "обл", "область", "р", "район",
}


def load_crm_dishes():
    try:
        resp = requests.get(f"{CRM_BASE_URL}/api/public/dishes", headers=_CRM_HEADERS, timeout=30)
        resp.raise_for_status()
        dishes = {}
        for item in resp.json().get("items", []):
            title = item.get("title", "").strip()
            dish_id = item.get("id")
            if title and dish_id:
                dishes[title.lower()] = {"id": dish_id, "weight": item.get("weight", 0) or 0}
        logger.info(f"Загружено блюд из CRM: {len(dishes)}")
        return dishes
    except Exception as e:
        logger.error(f"Не удалось загрузить блюда из CRM: {e}", exc_info=True)
        return {}


def load_delivery_times():
    return _DELIVERY_TIMES


def find_delivery_time_id(interval_str, delivery_times):
    if not interval_str or not delivery_times:
        return CRM_DELIVERY_TIME_ID
    if interval_str in delivery_times:
        tid = str(delivery_times[interval_str])
        logger.info(f"Интервал {interval_str!r} -> delivery_time_id={tid}")
        return tid
    m = re.search(r"(\d{1,2}:\d{2})[^\d]*(\d{1,2}:\d{2})", interval_str)
    if m:
        key = f"{m.group(1)}-{m.group(2)}"
        if key in delivery_times:
            tid = str(delivery_times[key])
            logger.info(f"Интервал {interval_str!r} -> delivery_time_id={tid}")
            return tid
    logger.warning(f"Интервал {interval_str!r} не найден в CRM, использую дефолтный id={CRM_DELIVERY_TIME_ID}")
    return CRM_DELIVERY_TIME_ID


def find_client_by_phone(phone):
    digits = re.sub(r"\D", "", phone)
    if not digits:
        return None
    try:
        resp = requests.get(
            f"{CRM_BASE_URL}/api/public/clients",
            headers=_CRM_HEADERS,
            params={"phone": digits},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        items = data.get("items") if isinstance(data, dict) else (data if isinstance(data, list) else [])
        if items:
            return items[0]
    except Exception as e:
        logger.warning(f"Поиск клиента по телефону {phone!r} не удался: {e}")
    return None


def _addr_tokens(addr):
    addr = addr.lower()
    addr = re.sub(r"[^\w\s]", " ", addr)
    return {w for w in addr.split() if w not in _ADDR_STOP and len(w) > 1}


def match_address(email_addr, client_addresses):
    """Возвращает address_id при нечётком совпадении (Jaccard >= ADDR_MATCH_THRESHOLD), иначе None."""
    email_tokens = _addr_tokens(email_addr)
    if not email_tokens:
        return None
    best_id, best_score = None, 0.0
    for addr_obj in client_addresses:
        addr_text = (
            addr_obj.get("address") or addr_obj.get("text") or addr_obj.get("title") or ""
        )
        addr_id = addr_obj.get("id")
        if not addr_text or not addr_id:
            continue
        crm_tokens = _addr_tokens(addr_text)
        if not crm_tokens:
            continue
        score = len(email_tokens & crm_tokens) / len(email_tokens | crm_tokens)
        if score > best_score:
            best_score, best_id = score, addr_id
    if best_score >= ADDR_MATCH_THRESHOLD:
        logger.info(f"Адрес найден (score={best_score:.2f}): {email_addr!r} -> address_id={best_id}")
        return best_id
    logger.info(f"Адрес не совпал (max_score={best_score:.2f}): {email_addr!r} — создаю новый")
    return None


def find_dish(dish_name, crm_dishes):
    name_lower = dish_name.strip().lower()
    if name_lower in crm_dishes:
        return crm_dishes[name_lower]
    name_words = set(name_lower.split())
    best_match, best_score = None, 0
    for crm_title, info in crm_dishes.items():
        score = len(name_words & set(crm_title.split()))
        if score > best_score and score >= 2:
            best_score = score
            best_match = info
    return best_match


def parse_order_email(html_body):
    soup = BeautifulSoup(html_body, "html.parser")
    text = soup.get_text("\n")
    order = {
        "name": "", "phone": "", "email": "", "address": "",
        "delivery_interval": "", "comment": "", "total_price": "0", "dishes": []
    }
    patterns = [
        ("name",              r"Покупатель:\s*(.+)"),
        ("phone",             r"Телефон покупателя:\s*(\S+)"),
        ("email",             r"Почта покупателя:\s*(\S+)"),
        ("address",           r"Адрес:\s*(.+)"),
        ("delivery_interval", r"Интервал доставки:\s*(.+)"),
        ("comment",           r"Комментарии к заказу:\s*(.+)"),
    ]
    for key, pattern in patterns:
        m = re.search(pattern, text)
        if m:
            order[key] = m.group(1).strip()
    m = re.search(r"Стоимость заказа[^\d]+([\d][\d\s ]*)", text)
    if m:
        order["total_price"] = re.sub(r"[^\d]", "", m.group(1)) or "0"
    table = soup.find("table")
    if table:
        for row in table.find_all("tr")[1:]:
            cells = row.find_all(["td", "th"])
            if len(cells) < 3:
                continue
            lines = [l.strip() for l in cells[0].get_text("\n", strip=True).split("\n") if l.strip()]
            if not lines:
                continue
            if lines[0] == "Название":
                continue
            if lines[0].startswith("Данные о заказе"):
                continue
            try:
                qty = int(cells[1].get_text(strip=True))
            except Exception:
                qty = 1
            try:
                price = int(re.sub(r"[^\d]", "", cells[2].get_text(strip=True))) if len(cells) > 2 else 0
            except Exception:
                price = 0
            sub_dishes = []
            for line in lines[1:]:
                for prefix in ["Суп:", "Второе блюдо:", "Салат:", "Напиток:", "Закуска:", "Десерт:", "Фрэш:"]:
                    if line.startswith(prefix):
                        name = line[len(prefix):].strip()
                        if name:
                            sub_dishes.append(name)
            if sub_dishes:
                sub_price = price // len(sub_dishes) if sub_dishes else 0
                for dish_name in sub_dishes:
                    order["dishes"].append({"title": dish_name, "count": qty, "price": sub_price})
            else:
                order["dishes"].append({"title": lines[0], "count": qty, "price": price})
    return order


def send_to_crm(order, crm_dishes, delivery_times):
    delivery_time_id = find_delivery_time_id(order["delivery_interval"], delivery_times)

    address_id = None
    if order["phone"] and order["address"]:
        client = find_client_by_phone(order["phone"])
        if client:
            addresses = client.get("addresses") or []
            if addresses:
                address_id = match_address(order["address"], addresses)

    url = f"{CRM_BASE_URL}/webApi/orderRequests/create"
    data = {
        "identifier": CRM_IDENTIFIER,
        "webApiKey": CRM_API_KEY,
        "name": order["name"] or "Клиент",
        "phone": order["phone"],
        "delivery_time_id": delivery_time_id,
        "order_source_id": CRM_ORDER_SOURCE_ID,
        "project_id": CRM_PROJECT_ID,
        "start_date": datetime.now().strftime("%Y-%m-%d"),
        "price": order["total_price"],
        "is_retail_order": "1",
        "count_person": "1",
    }
    if address_id:
        data["address_id"] = str(address_id)
    else:
        data["address_text"] = order["address"]

    for i, dish in enumerate(order["dishes"]):
        crm_dish = find_dish(dish["title"], crm_dishes)
        if crm_dish:
            data[f"retail_order_dishes[{i}][dish_id]"] = str(crm_dish["id"])
            data[f"retail_order_dishes[{i}][weight]"] = str(crm_dish["weight"])
            logger.info(f"    + {dish['title']} -> dish_id={crm_dish['id']}, weight={crm_dish['weight']}")
        else:
            data[f"retail_order_dishes[{i}][dish_title]"] = dish["title"]
            data[f"retail_order_dishes[{i}][weight]"] = "0"
            logger.info(f"    ? {dish['title']} -> не найдено в каталоге")
        data[f"retail_order_dishes[{i}][price]"] = str(dish["price"])
        data[f"retail_order_dishes[{i}][count]"] = str(dish["count"])

    logger.info(f"    [DEBUG] data['price'] = {data['price']!r}")
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                charset = part.get_content_charset() or "utf-8"
                return part.get_payload(decode=True).decode(charset, errors="replace")
    charset = msg.get_content_charset() or "utf-8"
    return msg.get_payload(decode=True).decode(charset, errors="replace")


def decode_subject(msg):
    parts = decode_header(msg.get("Subject", ""))
    result = []
    for p, c in parts:
        if isinstance(p, bytes):
            result.append(p.decode(c or "utf-8", errors="replace"))
        else:
            result.append(p)
    return "".join(result)


def load_processed_ids():
    try:
        with open(PROCESSED_IDS_FILE) as f:
            return set(json.load(f))
    except Exception:
        return set()


def save_processed_id(msg_key, processed_ids):
    processed_ids.add(msg_key)
    with open(PROCESSED_IDS_FILE, "w") as f:
        json.dump(list(processed_ids), f)


def process_emails():
    logger.info("Проверка почты...")
    crm_dishes = load_crm_dishes()
    delivery_times = load_delivery_times()
    _months_en = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    now = datetime.now()
    today = f"{now.day:02d}-{_months_en[now.month - 1]}-{now.year}"
    processed_ids = load_processed_ids()
    ok, err = 0, 0

    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)

        for folder in ["INBOX", "Spam"]:
            try:
                sel_status, _ = mail.select(folder, readonly=True)
                if sel_status != "OK":
                    logger.warning(f"Папка {folder!r} недоступна, пропускаю.")
                    continue

                status, ids = mail.search(None, f'SINCE "{today}"')
                if status != "OK" or not ids[0]:
                    logger.info(f"[{folder}] Новых писем нет.")
                    continue

                ids = ids[0].split()
                logger.info(f"[{folder}] Найдено писем за сегодня: {len(ids)}")

                for msg_id in ids:
                    try:
                        status, msg_data = mail.fetch(msg_id, "(BODY.PEEK[])")
                        if status != "OK":
                            continue
                        msg = email.message_from_bytes(msg_data[0][1])
                        subject = decode_subject(msg)
                        if "новый заказ" not in subject.lower():
                            continue
                        msg_key = (msg.get("Message-ID") or "").strip() or f"{subject}|{msg.get('Date', '')}"
                        if msg_key in processed_ids:
                            continue
                        logger.info(f"Обрабатываю: {subject}")
                        html_body = get_email_body(msg)
                        if not html_body:
                            continue
                        order = parse_order_email(html_body)
                        if not order["name"] and not order["phone"]:
                            logger.info("Не удалось извлечь данные, пропускаю.")
                            continue
                        logger.info(f"Клиент: {order['name']}, тел: {order['phone']}")
                        logger.info(f"Адрес: {order['address']}, сумма: {order['total_price']} руб.")
                        result = send_to_crm(order, crm_dishes, delivery_times)
                        crm_id = result.get("request", {}).get("id", "?")
                        logger.info(f"Заявка создана в CRM, ID: {crm_id}")
                        save_processed_id(msg_key, processed_ids)
                        ok += 1
                    except Exception as e:
                        logger.error(f"Ошибка при обработке письма {msg_id}: {e}", exc_info=True)
                        err += 1

            except Exception as e:
                logger.error(f"Ошибка при работе с папкой {folder}: {e}", exc_info=True)
                continue

    except Exception as e:
        logger.error(f"Критическая ошибка подключения к почте: {e}", exc_info=True)
        raise
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass

    logger.info(f"Итого: {ok} успешно, {err} ошибок.")


if __name__ == "__main__":
    logger.info("=== email_to_crm запущен ===")
    try:
        process_emails()
    except Exception as e:
        logger.error(f"Необработанное исключение: {e}", exc_info=True)
        sys.exit(1)
