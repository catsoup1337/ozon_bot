import telebot
from tqdm.contrib.telegram import tqdm, trange
import os
from dotenv import load_dotenv

from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import re
import requests
import time

import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

import yadisk

load_dotenv()
ua = UserAgent()
y = yadisk.YaDisk(token="AQAAAAAeeuFqAAeVCjRRWT3G8khEv1eCtEu6uY4")

TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN')
# CHAT_ID = os.getenv('TELEGRAM_CHAT_ID')
bot = telebot.TeleBot(TELEGRAM_TOKEN)
b = []

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    try:
        bot.reply_to(message, "Обрабатываю файл")
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        src_t = './documents/' + message.document.file_name
        with open(src_t, 'wb') as new_file:
            new_file.write(downloaded_file)
        name = message.document.file_name
        bot.reply_to(message, "Начал парсинг")
        read_data(name,src_t,file_info,message)
    except Exception as e:
        bot.reply_to(message, e)

def read_data(name, src_t, file_info, message):
    workbook_name2 = f'./documents/{name}'  
    wb2 = load_workbook(workbook_name2)
    ws2 = wb2.active
    sheet = wb2.get_sheet_by_name('Sheet1')
    df_init = len(pd.read_excel(f'./documents/{name}'))+2
    chat_id = message.chat.id
    counter = 1
    for i in trange(df_init, token=TELEGRAM_TOKEN, chat_id=chat_id):
        counter+=1
        res = sheet.cell(row=counter, column=1).value
        collect_data(name, src_t, file_info, counter, df_init, res, message)

def collect_data(name, src_t, file_info, counter, df_init, res, message):
    res1 = res
    res = res.replace(' ', '+')
    url = f'https://www.ozon.ru/search/?from_global=true&text={res}'
    response = requests.get(
                url=url,
                headers={'user-agent': f'{ua.random}',
                'Cookie':'_ga_JNVTMNXQ6F=GS1.1.1640426535.11.1.1640427081.19; tmr_reqNum=230; _dc_gtm_UA-37420525-1=1; _fbp=fb.1.1640245791587.379020311; _ga=GA1.2.927695446.1640245781; _gid=GA1.2.1342864285.1640245783; userId=86541350; __exponea_etc__=4d6c5ea1-34a2-4095-8d7e-b73d96ae5c5e; cto_bundle=CfI1ll9HbmRlZDNqZVVxY1ZHOFdXNEJPZ1BDVkZacURNdDZCd2hvd3RKVTVZaEloYjlpRjFxcHY4RlpERTliRCUyRnNZT09RSGV4bmt2UlVFdmtZJTJCMlJYVkRoVHpuJTJCJTJCSk82YjFjVXZKcjBYJTJGU2pPMTglM0Q; tmr_detect=1%7C1640426975278; __exponea_time2__=0.11290812492370605; tmr_lvid=287c4c516bef2174f533ac42f0a0f1da; tmr_lvidTS=1639525512899; is_adult_confirmed=true; incap_ses_379_1101384=lT7OPWvshVxvwwOlHntCBSHsxmEAAAAAdA8KcPYeoZQQi1ZBIGvZtQ==; __Secure-ab-group=79; __Secure-access-token=3.86541350.RG5vuMilQ9yzbV8tIYZQtQ.79.l8cMBQAAAABhxfKnCaZmq6N3ZWKrNzk5OTcxNDAxNzQAgJCg.20211224161743.20211225120208.uVg4GqBI79dOdXnE4aANIgyRmQYRIuVwTc4rlBf_2EE; __Secure-refresh-token=3.86541350.RG5vuMilQ9yzbV8tIYZQtQ.79.l8cMBQAAAABhxfKnCaZmq6N3ZWKrNzk5OTcxNDAxNzQAgJCg.20211224161743.20211225120208.GAeBJpoAXPTSCTTkRbm7dvCJafZBtAsI1RAVoQfuuWQ; __Secure-user-id=86541350; incap_ses_585_1101384=CrVILF8aoTjhHTR8JVceCCDsxmEAAAAAKIwfvHqBoIgMLHOeJ0N5eg==; nlbi_1101384=FymcANTOD3yrXO/nK8plmQAAAACdjpmwjeRed1mk5/5xViDD; xcid=b685b6ecd560a6efaf5a6b6204bc20f2; cnt_of_orders=0; isBuyer=0; _gcl_au=1.1.1311607219.1640245779; __Secure-ext_xcid=2cbaa67316c6ea12508f03aa9ca90041; visid_incap_1101384=apJ7+ksQQii0kXoTFnZAzYAsuWEAAAAAQUIPAAAAAAAkQumycxPLkRpl0eIoWpAd'}
            )
    src = response.text
    try:
        soup_split = BeautifulSoup(src, 'lxml')
        script = str(soup_split.find_all("script", attrs= {"type": "application/javascript"})[1])
        splited = script.split('>')[1].split('"')[1].replace("\/",'/').split('category_was_predicted=true')[0]
        url1 = f'https://www.ozon.ru{splited}category_was_predicted=true&from_global=true&text={res}'
        response1 = requests.get(
                url=url1,
                headers={'user-agent': f'{ua.random}',
                'Cookie':'_ga_JNVTMNXQ6F=GS1.1.1640426535.11.1.1640427081.19; tmr_reqNum=230; _dc_gtm_UA-37420525-1=1; _fbp=fb.1.1640245791587.379020311; _ga=GA1.2.927695446.1640245781; _gid=GA1.2.1342864285.1640245783; userId=86541350; __exponea_etc__=4d6c5ea1-34a2-4095-8d7e-b73d96ae5c5e; cto_bundle=CfI1ll9HbmRlZDNqZVVxY1ZHOFdXNEJPZ1BDVkZacURNdDZCd2hvd3RKVTVZaEloYjlpRjFxcHY4RlpERTliRCUyRnNZT09RSGV4bmt2UlVFdmtZJTJCMlJYVkRoVHpuJTJCJTJCSk82YjFjVXZKcjBYJTJGU2pPMTglM0Q; tmr_detect=1%7C1640426975278; __exponea_time2__=0.11290812492370605; tmr_lvid=287c4c516bef2174f533ac42f0a0f1da; tmr_lvidTS=1639525512899; is_adult_confirmed=true; incap_ses_379_1101384=lT7OPWvshVxvwwOlHntCBSHsxmEAAAAAdA8KcPYeoZQQi1ZBIGvZtQ==; __Secure-ab-group=79; __Secure-access-token=3.86541350.RG5vuMilQ9yzbV8tIYZQtQ.79.l8cMBQAAAABhxfKnCaZmq6N3ZWKrNzk5OTcxNDAxNzQAgJCg.20211224161743.20211225120208.uVg4GqBI79dOdXnE4aANIgyRmQYRIuVwTc4rlBf_2EE; __Secure-refresh-token=3.86541350.RG5vuMilQ9yzbV8tIYZQtQ.79.l8cMBQAAAABhxfKnCaZmq6N3ZWKrNzk5OTcxNDAxNzQAgJCg.20211224161743.20211225120208.GAeBJpoAXPTSCTTkRbm7dvCJafZBtAsI1RAVoQfuuWQ; __Secure-user-id=86541350; incap_ses_585_1101384=CrVILF8aoTjhHTR8JVceCCDsxmEAAAAAKIwfvHqBoIgMLHOeJ0N5eg==; nlbi_1101384=FymcANTOD3yrXO/nK8plmQAAAACdjpmwjeRed1mk5/5xViDD; xcid=b685b6ecd560a6efaf5a6b6204bc20f2; cnt_of_orders=0; isBuyer=0; _gcl_au=1.1.1311607219.1640245779; __Secure-ext_xcid=2cbaa67316c6ea12508f03aa9ca90041; visid_incap_1101384=apJ7+ksQQii0kXoTFnZAzYAsuWEAAAAAQUIPAAAAAAAkQumycxPLkRpl0eIoWpAd'}
            )
        src1 = response1.text
        soup1 = BeautifulSoup(src1, 'lxml')
        try:
            search = soup1.find_all(class_="b6r7")[0].text.replace(res1,'')
            qty = ''.join(re.findall('\d', search))
            b.append(qty)
        except:
            search = soup1.find_all(class_="b3a1")[0].text.replace(res1,'')
            qty = ''.join(re.findall('\d', search))
            b.append(qty)
    except:
        try:
            soup_usualy = BeautifulSoup(src, 'lxml')
            search = soup_usualy.find_all(class_="b6r7")[0].text.replace(res1,'')
            qty = ''.join(re.findall('\d', search))
            b.append(qty)
        except:
            qty = 0
            b.append(qty)
    if counter == df_init:
        data = {'Data': b}
        saver(name, src_t, file_info, data, message)

def saver(name, src_t, file_info, data, message):
    workbook_name1 = f'./documents/{name}' 
    df = pd.DataFrame(data)
    writer = pd.ExcelWriter(workbook_name1, engine='xlsxwriter')
    df.to_excel(writer,startcol = 2, sheet_name='Sheet1',index=False, header=False)
    writer.save()
    y.upload(src_t, f'{file_info.file_path}')
    d_link = y.get_download_link(file_info.file_path)
    bot.reply_to(message, f"Всё готово, вот ссылка на скачивание {d_link}")

bot.infinity_polling()
