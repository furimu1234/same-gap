import requests
import json
import pprint

import asyncio
from aiohttp import ClientSession

from config import SCRIPTID, BOOKID

async def main(data="test sheet"):
    
    payload = {
        "gmail": "example@gmail.com",
        "type": "create",
        "filename": "test sheet",
        "sheetname": "基本設定",
        #"bookid": BOOKID,
        "row": 3,
        "col": 3,
        #"value": "test OK!\nOK!!!!!!!!!!!!"
    
    }

    url = f"https://script.google.com/macros/s/{SCRIPTID}/exec"

    headers = {
        'Content-Type': 'application/json',
    }

    async with ClientSession() as session:
        async with session.post(url, data=json.dumps(payload), headers=headers) as resp:
            print(resp.status)
            print(await resp.text())


if __name__ == "__main__":
    # postしたいデータを渡す

    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())