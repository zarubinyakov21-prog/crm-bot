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
PROCESSED_FOLDER = "Olympic_Processed"
SUBJECT_FILTER = "заказ с сайта olympickitchen"
# =====================


def parse_order_email(body):
    soup = BeautifulSoup(body, "html.parser")
    text = soup.get_text("\n")
    order = {
        "name": "", "phone": "", "address": "",
        "price": "0", "program": "", "comment": "",
        "duration": "", "no_weekend": "",
    }
    patterns = [
        ("address",    r"Адрес:\s*(.+)"),
        ("name",       r"Имя:\s*(.+)"),
        ("phone",      r"Телефон:\s*(.+)"),
        ("price",      r"Цена:\s*([\d]+)"),
        ("program",    r"Программа:\s*(.+)"),
        ("comment",    r"Сообщение:\s*(.+)"),
        ("duration",   r"Срок:\s*(.+)"),
        ("no_weekend", r"Без доставки на выходные:\s*(.+)"),
    ]
    for key, pattern in patterns:
        m = re.search(pattern, text)
        if m:
            value = m.group(1).strip()
            if value.lower() != "(пусто)":
                order[key] = value
    return order


def send_to_crm(order):
    url = f"{CRM_BASE_URL}/webApi/orderRequests/create"
    comment_parts = [f"Программа: {order['program']}"] if order["program"] else []
    if order["duration"]:
        comment_parts.append(f"Срок: {order['duration']}")
    if order["no_weekend"]:
        comment_parts.append(f"Без доставки на выходные: {order['no_weekend']}")
    if order["comment"]:
        comment_parts.append(f"Комментарий: {order['comment']}")
    comment_parts.append("Источник: сайт olympickitchen.ru")
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
        "price": order["price"],
        "count_person": "1",
    }
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
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
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
    print(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] Проверка olympic почты...")
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
        if SUBJECT_FILTER not in subject.lower():
            continue
        print(f"\n  Обрабатываю: {subject}")
        body = get_email_body(msg)
        if not body:
            continue
        order = parse_order_email(body)
        if not order["name"] and not order["phone"]:
            print("    Не удалось извлечь данные, пропускаю.")
            continue
        print(f"    Клиент: {order['name']}, тел: {order['phone']}")
        print(f"    Адрес: {order['address']}, программа: {order['program']}, сумма: {order['price']} руб.")
        try:
            result = send_to_crm(order)
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
