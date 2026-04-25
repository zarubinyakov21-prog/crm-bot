import imaplib, email, re, requests, os
from email.header import decode_header
from bs4 import BeautifulSoup
from datetime import datetime

# ===== НАСТРОЙКИ =====
IMAP_HOST = "imap.yandex.ru"
IMAP_PORT = 993
EMAIL_LOGIN = os.environ.get("EMAIL_LOGIN", "zakaz@olympickitchen.ru")
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
CRM_BASE_URL = "https://crm.private-crm.ru"
CRM_IDENTIFIER = "ok"
CRM_API_KEY = os.environ["CRM_API_KEY"]
CRM_PROJECT_ID = "1"
CRM_DELIVERY_TIME_ID = "1"
CRM_ORDER_SOURCE_ID = "1"
PROCESSED_FOLDER = "CRM_Processed"
# =====================


def load_crm_dishes():
    url = f"{CRM_BASE_URL}/api/public/dishes"
    headers = {"Identifier": CRM_IDENTIFIER, "Application-key": CRM_API_KEY}
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        dishes = {}
        for item in resp.json().get("items", []):
            title = item.get("title", "").strip()
            dish_id = item.get("id")
            if title and dish_id:
                dishes[title.lower()] = {"id": dish_id, "weight": item.get("weight", 0) or 0}
        print(f"  Загружено блюд из CRM: {len(dishes)}")
        return dishes
    except Exception as e:
        print(f"  Не удалось загрузить блюда: {e}")
        return {}


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
        ("name", r"Покупатель:\s*(.+)"),
        ("phone", r"Телефон покупателя:\s*(\S+)"),
        ("email", r"Почта покупателя:\s*(\S+)"),
        ("address", r"Адрес:\s*(.+)"),
        ("delivery_interval", r"Интервал доставки:\s*(.+)"),
        ("comment", r"Комментарии к заказу:\s*(.+)"),
    ]
    for key, pattern in patterns:
        m = re.search(pattern, text)
        if m:
            order[key] = m.group(1).strip()
    m = re.search(r"Стоимость заказа[^\d]+([\d][\d\s ]*)", text)
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


def send_to_crm(order, crm_dishes):
    url = f"{CRM_BASE_URL}/webApi/orderRequests/create"
    comment_parts = []
    if order["delivery_interval"]:
        comment_parts.append(f"Интервал доставки: {order['delivery_interval']}")
    if order["comment"]:
        comment_parts.append(f"Комментарий: {order['comment']}")
    if order["email"]:
        comment_parts.append(f"Email: {order['email']}")
    comment_parts.append("Источник: сайт pp-obed.ru")
    data = {
        "identifier": CRM_IDENTIFIER,
        "webApiKey": CRM_API_KEY,
        "name": order["name"] or "Клиент",
        "phone": order["phone"],
        "address_text": order["address"],
        "comment": "\n".join(comment_parts),
        "delivery_time_id": CRM_DELIVERY_TIME_ID,
        "order_source_id": CRM_ORDER_SOURCE_ID,
        "project_id": CRM_PROJECT_ID,
        "start_date": datetime.now().strftime("%Y-%m-%d"),
        "price": order["total_price"],
        "is_retail_order": "1",
        "count_person": "1",
    }
    for i, dish in enumerate(order["dishes"]):
        crm_dish = find_dish(dish["title"], crm_dishes)
        if crm_dish:
            data[f"retail_order_dishes[{i}][dish_id]"] = str(crm_dish["id"])
            data[f"retail_order_dishes[{i}][weight]"] = str(crm_dish["weight"])
            print(f"    + {dish['title']} -> dish_id={crm_dish['id']}, weight={crm_dish['weight']}")
        else:
            data[f"retail_order_dishes[{i}][dish_title]"] = dish["title"]
            data[f"retail_order_dishes[{i}][weight]"] = "0"
            print(f"    ? {dish['title']} -> не найдено в каталоге")
        data[f"retail_order_dishes[{i}][price]"] = str(dish["price"])
        data[f"retail_order_dishes[{i}][count]"] = str(dish["count"])
    print(f"    [DEBUG] data['price'] = {data['price']!r}")
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


def process_emails():
    print(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] Проверка почты...")
    crm_dishes = load_crm_dishes()
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)
    mail.select("INBOX")
    today = datetime.now().strftime("%d-%b-%Y")
    status, ids = mail.search(None, f'(UNSEEN SINCE "{today}")')
    if status != "OK" or not ids[0]:
        print("  Новых заказов нет.")
        mail.logout()
        return
    ids = ids[0].split()
    print(f"  Найдено писем сегодня: {len(ids)}")
    mail.create(PROCESSED_FOLDER)
    ok, err = 0, 0
    for msg_id in ids:
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        subject = decode_subject(msg)
        if "новый заказ" not in subject.lower():
            continue
        print(f"\n  Обрабатываю: {subject}")
        html_body = get_email_body(msg)
        if not html_body:
            continue
        order = parse_order_email(html_body)
        if not order["name"] and not order["phone"]:
            print("    Не удалось извлечь данные, пропускаю.")
            continue
        print(f"    Клиент: {order['name']}, тел: {order['phone']}")
        print(f"    Адрес: {order['address']}, сумма: {order['total_price']} руб.")
        try:
            result = send_to_crm(order, crm_dishes)
            crm_id = result.get("request", {}).get("id", "?")
            print(f"    Заявка создана в CRM, ID: {crm_id}")
            mail.store(msg_id, "+FLAGS", "\\Seen")
            mail.copy(msg_id, PROCESSED_FOLDER)
            mail.store(msg_id, "+FLAGS", "\\Deleted")
            ok += 1
        except Exception as e:
            print(f"    Ошибка: {e}")
            err += 1
    mail.expunge()
    mail.logout()
    print(f"\n  Итого: {ok} успешно, {err} ошибок.")


if __name__ == "__main__":
    process_emails()
