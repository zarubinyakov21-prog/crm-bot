import os, json, base64, logging, requests
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes
import anthropic

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
ANTHROPIC_API_KEY  = os.environ["ANTHROPIC_API_KEY"]
CRM_BASE_URL       = "https://crm.private-crm.ru"
CRM_IDENTIFIER     = "ok"
CRM_API_KEY        = os.environ["CRM_API_KEY"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


PROMPT = """Ты — система распознавания товарных накладных. На фото накладная.
Извлеки данные и верни ТОЛЬКО JSON без лишнего текста:
{
  "supplier": "название поставщика или null",
  "invoice_number": "номер накладной или null",
  "date": "дата в формате YYYY-MM-DD или null",
  "items": [
    {"title": "название товара", "qty": число, "unit": "ед.изм.", "price": число, "total": число}
  ],
  "total_sum": итоговая сумма числом
}
Если поле не читается — ставь null."""


def recognize_invoice(image_bytes: bytes) -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    b64 = base64.standard_b64encode(image_bytes).decode()
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": PROMPT},
            ],
        }],
    )
    text = msg.content[0].text.strip()
    # вырезаем JSON если Claude обернул его в ```
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)


def send_to_crm(invoice: dict) -> dict:
    url = f"{CRM_BASE_URL}/webApi/warehouseInvoice/create"
    data = {
        "identifier": CRM_IDENTIFIER,
        "webApiKey": CRM_API_KEY,
        "supplier": invoice.get("supplier") or "",
        "invoice_number": invoice.get("invoice_number") or "",
        "date": invoice.get("date") or "",
        "total_sum": str(invoice.get("total_sum") or 0),
    }
    for i, item in enumerate(invoice.get("items", [])):
        data[f"items[{i}][title]"]  = item.get("title", "")
        data[f"items[{i}][qty]"]    = str(item.get("qty", 0))
        data[f"items[{i}][unit]"]   = item.get("unit", "шт")
        data[f"items[{i}][price]"]  = str(item.get("price", 0))
        data[f"items[{i}][total]"]  = str(item.get("total", 0))
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    await msg.reply_text("Распознаю накладную...")
    try:
        photo = await msg.photo[-1].get_file()
        image_bytes = await photo.download_as_bytearray()
        invoice = recognize_invoice(bytes(image_bytes))
        logging.info("Распознано: %s", invoice)

        items_text = "\n".join(
            f"  • {it['title']}: {it['qty']} {it.get('unit','шт')} × {it.get('price',0)} = {it.get('total',0)} руб."
            for it in invoice.get("items", [])
        )
        summary = (
            f"Поставщик: {invoice.get('supplier') or '—'}\n"
            f"Номер: {invoice.get('invoice_number') or '—'}\n"
            f"Дата: {invoice.get('date') or '—'}\n"
            f"Товары:\n{items_text}\n"
            f"Итого: {invoice.get('total_sum') or 0} руб."
        )

        try:
            result = send_to_crm(invoice)
            crm_id = result.get("id") or result.get("invoice", {}).get("id", "?")
            await msg.reply_text(f"Накладная загружена в CRM (ID: {crm_id})\n\n{summary}")
        except Exception as e:
            logging.warning("CRM ошибка: %s", e)
            await msg.reply_text(
                f"Данные распознаны, но загрузка в CRM не удалась ({e}).\n\n{summary}"
            )
    except Exception as e:
        logging.error("Ошибка: %s", e)
        await msg.reply_text(f"Не удалось распознать накладную: {e}")


def main():
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    logging.info("Бот запущен")
    app.run_polling()


if __name__ == "__main__":
    main()
