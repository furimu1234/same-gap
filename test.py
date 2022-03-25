import asyncio
from aiohttp import ClientSession

from config import SCRIPTID, BOOKID


async def main():

    payload = {
        "gmail": "example@gmail.com",
        "type": "get",
        "filename": "754680802578792498",
        "sheetname": "基本設定",
        "bookid": BOOKID,
        "range": "B5:H5",
        # "value": "test OK!\nOK!!!!!!!!!!!!"
    }

    url = f"https://script.google.com/macros/s/{SCRIPTID}/exec"

    url = "https://script.google.com/macros/s/AKfycbzIfXZqGcSS22-3o4ceySO_fyy2yEsHwEAh4IKxcZRqBxdS5gLcL7zn2CaT4RILZbHg/exec"
    headers = {
        "Content-Type": "application/json",
    }

    async with ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            print(resp.status)
            print((await resp.text()))


if __name__ == "__main__":

    # asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
