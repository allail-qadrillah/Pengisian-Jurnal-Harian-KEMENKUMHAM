from app import BOT
import logging

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")

bot = BOT(server="local")
bot.start()

