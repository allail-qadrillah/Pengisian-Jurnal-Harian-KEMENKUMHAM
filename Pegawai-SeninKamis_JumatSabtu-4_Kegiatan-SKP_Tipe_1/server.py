from app import BOT
import logging

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

bot = BOT(server="lambda")

def main_task(event=None, context=None):
  bot.start()

    