from telegram import Bot
from telegram.ext import ApplicationBuilder
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import asyncio
import os
from PIL import Image

BOT_TOKEN = '7522224142:AAHKUUZFcW-PD3uob9LJcFkCVguhI30feaQ'
CHANNEL_ID = '@programm_123'
OUTPUT_DIR = '../Task_12/'

class ImageHandler(FileSystemEventHandler):
    def __init__(self, bot, loop):
        self.bot = bot
        self.loop = loop
        self.processed_files = set()  # Track processed files

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
            # Check if the file has already been processed
            if event.src_path not in self.processed_files:
                # Schedule the coroutine in the main event loop
                self.loop.call_soon_threadsafe(asyncio.create_task, self.post_image_to_telegram(event.src_path))

    async def post_image_to_telegram(self, image_path, retries=3):
        try:
            await asyncio.sleep(1)  # Wait a bit to ensure the file is fully written

            if not os.path.exists(image_path) or not os.path.isfile(image_path):
                print(f"File {image_path} does not exist or is not a file.")
                return

            # Optimize the image if desired
            optimize_image(image_path)

            with open(image_path, 'rb') as image:
                await self.bot.send_document(chat_id=CHANNEL_ID, document=image)
                print(f"Posted {image_path} to Telegram as document.")

        except Exception as e:
            if retries > 0:
                print(f"Failed to post {image_path}. Retrying... ({retries} attempts left)")
                await asyncio.sleep(2)  # Wait a bit before retrying
                await self.post_image_to_telegram(image_path, retries - 1)
            else:
                print(f"Failed to post {image_path} to Telegram: {e}")

def optimize_image(image_path):
    with Image.open(image_path) as img:
        img = img.convert('RGB')
        img.save(image_path, 'JPEG', quality=95)  # Adjust quality as needed

async def send_existing_images(bot, loop, processed_files):
    for filename in os.listdir(OUTPUT_DIR):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
            image_path = os.path.join(OUTPUT_DIR, filename)
            if os.path.isfile(image_path):
                # Check if the file has already been processed
                if image_path not in processed_files:
                    # Schedule the task to post existing images
                    loop.call_soon_threadsafe(asyncio.create_task, ImageHandler(bot, loop).post_image_to_telegram(image_path))

async def main():
    application = ApplicationBuilder().token(BOT_TOKEN).build()
    bot = application.bot
    loop = asyncio.get_running_loop()

    # Initialize ImageHandler with an empty set of processed files
    handler = ImageHandler(bot, loop)

    # Send existing images first
    await send_existing_images(bot, loop, handler.processed_files)

    # Set up monitoring for new images
    observer = Observer()
    observer.schedule(handler, path=OUTPUT_DIR, recursive=False)
    observer.start()

    print(f"Monitoring directory {OUTPUT_DIR} for new images...")

    try:
        while True:
            await asyncio.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    # Run the main function in an asyncio event loop
    asyncio.run(main())
