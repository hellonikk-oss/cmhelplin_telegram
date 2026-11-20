
import os
import asyncio
from xmlrpc import client
from dotenv import load_dotenv
from telethon import TelegramClient

load_dotenv()

API_ID = int(os.getenv("API_ID"))
API_HASH = os.getenv("API_HASH")
PHONE = os.getenv("PHONE_NUMBER")

async def main():

    async with TelegramClient(
        "my_session",
        API_ID,
        API_HASH
    ) as client:

        # Friend's Telegram username
        friend = "nikhil_claw_bot"  # Replace with your friend's phone number or username

        # Excel file path
        folder_name = 'cmhelpline'
        file1_1 = os.path.join(folder_name, 'L-1.xlsx') 
        file1_2 = os.path.join(folder_name, 'high_level.xlsx')
        files = [
            file1_1,
            file1_2
        ]
        # Send Excel file
        await client.send_file(
            friend,
            files,
            caption="Here is your Excel report."
        )

        print("Excel file sent successfully!")

asyncio.run(main())

