import requests
import os

class TelegramLogger:
    def __init__(self, token=None, chat_id=None):
        self.token = token or os.getenv("TELEGRAM_BOT_TOKEN")
        self.chat_id = chat_id or os.getenv("TELEGRAM_CHAT_ID")
        self.base_url = f"https://api.telegram.org/bot{self.token}/sendMessage"

    def send_message(self, message, parse_mode=None):
        if not self.token or not self.chat_id:
            print("Telegram token or chat_id not set. Skipping message.")
            return

        try:
            payload = {
                "chat_id": self.chat_id,
                "text": message
            }
            if parse_mode:
                payload["parse_mode"] = parse_mode
            response = requests.post(self.base_url, data=payload)
            response.raise_for_status()
        except Exception as e:
            print(f"Failed to send Telegram message: {e}")
