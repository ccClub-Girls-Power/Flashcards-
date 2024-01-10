# 載入LineBot所需要的套件
from flask import Flask, request, abort

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import *

app = Flask(__name__)

# 必須放上自己的Channel Access Token
line_bot_api = LineBotApi(
    'Yxy/j1WMkLcW+qV2OJb1YIjx9V4sGCcUmxlP8FTtHiEkr/P5TiaOwocFqWnpeS1G1D7S/QO/Qibsd6P3u2pf7An14bzy/nb10NNfFtKgyIITsFNIB4Wl9o4xNzHk7Cgk+hab356oMaMA4gkt7tu26wdB04t89/1O/w1cDnyilFU=')
# 必須放上自己的Channel Secret
handler = WebhookHandler('9b53acd27ad8b0b02b9f379e21f1305b')


# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'


# 訊息傳遞區塊
##### 程式編輯都在這個function #####
# 載入必要的套件
from datetime import datetime
from linebot.models import FlexSendMessage, TextSendMessage, BubbleContainer
import pygsheets
from linebot.models import MessageEvent, TextMessage
import pandas as pd
import pytz  # 指定時區
import requests
from bs4 import BeautifulSoup


# 函數: 取得所有工作表
def get_all_worksheets(spreadsheet_url, service_file_path, user_id, deck_name):
    gc = pygsheets.authorize(service_file=service_file_path)
    worksheet_title = f'{user_id}_{deck_name}'
    try:
        worksheet = gc.open_by_url(spreadsheet_url).worksheet_by_title(worksheet_title)
        return worksheet
    except pygsheets.exceptions.WorksheetNotFound:
        return None


# 函數: 建立新工作表
def create_new_worksheet(user_id, deck_name, service_file_path, spreadsheet_url):
    gc = pygsheets.authorize(service_file=service_file_path)
    spreadsheet = gc.open_by_url(spreadsheet_url)
    new_worksheet_title = f'{user_id}_{deck_name}'
    new_worksheet = spreadsheet.add_worksheet(new_worksheet_title)
    return new_worksheet


# 儲存卡片失敗提示
class SaveCardError(Exception):
    pass


# 函數: 儲存閃卡卡片內容至工作表
def save_card_content_to_sheet(current_time, sheet_title, content_list, service_file_path, spreadsheet_url):
    try:
        # 開啟 Google Sheets
        gc = pygsheets.authorize(service_file=service_file_path)
        spreadsheet = gc.open_by_url(spreadsheet_url)
        # 選擇對應的工作表
        worksheet = spreadsheet.worksheet_by_title(sheet_title)
        # 將卡片內容轉換為字典
        data = {'新增時間': [current_time], "卡片正面": [content_list[0]], "卡片背面": [content_list[1]]}
        # 將字典轉換為 DataFrame
        df = pd.DataFrame(data)
        # 將 DataFrame 寫入 Google Sheets
        worksheet.set_dataframe(df, start='A1')  # 將 DataFrame 從第一行的第一列開始寫入
        return True  # 儲存成功，返回 True
    except Exception as e:
        # 如果儲存失敗，引發自定義的 SaveCardError 並帶上錯誤訊息
        raise SaveCardError("卡片儲存失敗，請稍後再試。")


# 函數: 儲存閃卡卡片內容至工作表(插入現有行數後面）
def insert_card_content_to_sheet(current_time, deck_name, card_contents, service_file_path, spreadsheet_url):
    try:
        # 開啟 Google Sheets
        gc = pygsheets.authorize(service_file=service_file_path)
        spreadsheet = gc.open_by_url(spreadsheet_url)
        # 選擇對應的工作表，如果不存在會自動建立
        worksheet = spreadsheet.worksheet_by_title(deck_name)
        # 讀取工作表內容並轉換為 Pandas DataFrame
        df = worksheet.get_as_df()
        # 獲取 DataFrame 的形狀
        num_rows = df.shape[0]
        # 新數據追加到下一行
        start_row = num_rows + 1
        # 定義要插入的資料，包含新增時間
        data = [current_time, card_contents[0], card_contents[1]]
        # 在指定的行數插入新數據
        worksheet.insert_rows(start_row, values=[data])
        return start_row
    except Exception as e:
        # 如果儲存失敗，引發自定義的 SaveCardError 並帶上錯誤訊息
        raise SaveCardError("卡片儲存失敗，請稍後再試。")

# 函數: 儲存單字卡片內容至工作表
def save_word_card_content_to_sheet(current_time, sheet_title, content_list, service_file_path, spreadsheet_url):
    try:
        # 開啟 Google Sheets
        gc = pygsheets.authorize(service_file=service_file_path)
        spreadsheet = gc.open_by_url(spreadsheet_url)
        worksheet = spreadsheet.worksheet_by_title(sheet_title)
        # 將卡片內容轉換為字典
        data = {'新增時間': [current_time], "單字": [content_list[0]], "詞性": [content_list[1]],
                "中文": [content_list[2]], "例句": [content_list[3]], "筆記": [content_list[4]]}
        # 將字典轉換為 DataFrame
        df = pd.DataFrame(data)
        # 將 DataFrame 寫入 Google Sheets
        worksheet.set_dataframe(df, start='A1')  # 將 DataFrame 從第一行的第一列開始寫入
        return True  # 儲存成功的情況下返回 True
    except Exception as e:
        # 如果儲存失敗，引發自定義的 SaveCardError 並帶上錯誤訊息
        raise SaveCardError("卡片儲存失敗，請稍後再試。")


# 函數: 儲存單字卡片內容至工作表(插入現有行數後面）
def insert_word_card_content_to_sheet(current_time, deck_name, card_contents, service_file_path, spreadsheet_url):
    try:
        # 開啟 Google Sheets
        gc = pygsheets.authorize(service_file=service_file_path)
        spreadsheet = gc.open_by_url(spreadsheet_url)
        # 選擇對應的工作表
        worksheet = spreadsheet.worksheet_by_title(deck_name)
        # 讀取工作表內容並轉換為 Pandas DataFrame
        df = worksheet.get_as_df()
        # 獲取 DataFrame 的形狀
        num_rows = df.shape[0]
        # 新數據追加到下一行
        start_row = num_rows + 1
        # 定義要插入的資料，包含新增時間
        data = [current_time] + card_contents
        # 在指定的行數插入新數據
        worksheet.insert_rows(start_row, values=[data])
        return start_row
    except Exception as e:
        # 如果儲存失敗，引發自定義的 SaveCardError 並帶上錯誤訊息
        raise SaveCardError("卡片儲存失敗，請稍後再試。")


# 函數: 查單字爬蟲
def lookup_word(word):
    url = f"https://dictionary.cambridge.org/zht/%E8%A9%9E%E5%85%B8/%E8%8B%B1%E8%AA%9E-%E6%BC%A2%E8%AA%9E-%E7%B9%81%E9%AB%94/{word}"

    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")

    entries = soup.select(".entry-body__el")

    pos_list = []
    example_list = []
    us_pronunciation_url = None
    uk_pronunciation_url = None

    for i, entry in enumerate(entries, 1):
        pos_element = entry.select_one(".pos.dpos")
        if pos_element:
            pos = pos_element.get_text(strip=True)
        else:
            pos = "N/A"

        chinese_definition_element = entry.select_one(".trans.dtrans.dtrans-se")
        if chinese_definition_element:
            chinese_definition = chinese_definition_element.text.strip()
        else:
            chinese_definition = "N/A"

        pos_dic = {"pos": pos, "chinese_definition": chinese_definition}
        pos_list.append(pos_dic)

        example_item = (
                f"{i}. {pos}" + "\n" + f"中文解釋：{chinese_definition}" + "\n" + "例句："
        )

        examples = entry.select(".def-body.ddef_b .examp.dexamp")

        if not examples:
            example_item += "\n無"
        else:
            for j, example in enumerate(examples, 1):
                eng_example = example.select_one(".eg.deg").text.strip()
                chinese_example_element = example.select_one(".trans.dtrans.dtrans-se")
                if chinese_example_element:
                    chinese_example = chinese_example_element.text.strip()
                else:
                    chinese_example = "N/A"

                example_item += "\n" + f"{eng_example}\n{chinese_example}"
                if j == 2:
                    break

        example_list.append(example_item)

    # 找到英式發音的音檔
    uk_audio_element = soup.find('audio', {'id': 'audio1'})
    uk_source_element = uk_audio_element.find('source', {'type': 'audio/mpeg'})
    uk_pronunciation_url = "https://dictionary.cambridge.org" + uk_source_element['src'] if uk_source_element else None

    # 找到美式發音的音檔
    us_audio_element = soup.find('audio', {'id': 'audio2'})
    us_source_element = us_audio_element.find('source', {'type': 'audio/mpeg'})
    us_pronunciation_url = "https://dictionary.cambridge.org" + us_source_element['src'] if us_source_element else None

    return pos_list, example_list, us_pronunciation_url, uk_pronunciation_url


# 函數：儲存字典卡片內容至工作表（建立新工作表+插入標題欄）
def searching_word_to_sheet(current_time, service_file_path, spreadsheet_url, sheet_title, word, pos_list, example_list,
                            us_pronunciation_url, uk_pronunciation_url):
    try:
        # 開啟 Google Sheets
        gc = pygsheets.authorize(service_file=service_file_path)
        spreadsheet = gc.open_by_url(spreadsheet_url)
        # 新增工作表
        worksheet = spreadsheet.add_worksheet(sheet_title)
        df = pd.DataFrame({
            '新增時間': [current_time],
            '單字': [word],
            '詞性': [','.join(pos['pos'] for pos in pos_list)],
            '中文': [','.join(pos['chinese_definition'] for pos in pos_list)],
            '例句': [','.join(example for example in example_list)],
            'US Pronunciation': [us_pronunciation_url] if us_pronunciation_url else [''],
            'UK Pronunciation': [uk_pronunciation_url] if uk_pronunciation_url else ['']
        })
        # 插入資料到 Google Sheets
        worksheet.set_dataframe(df, start='A1')  # 從 A1 開始插入 DataFrame 到工作表
        return True  # 儲存成功的情況下返回 True
    except Exception as e:
        # 如果儲存失敗，引發自定義的 SaveCardError 並帶上錯誤訊息
        raise SaveCardError("卡片儲存失敗，請稍後再試。")


# 函數：儲存字典卡片內容至現有工作表（插入現有行數後面）
def searching_word_to_existing_sheet(current_time, service_file_path, spreadsheet_url, sheet_title, word, pos_list,
                                     example_list, us_pronunciation_url, uk_pronunciation_url):
    try:
        # 開啟 Google Sheets
        gc = pygsheets.authorize(service_file=service_file_path)
        spreadsheet = gc.open_by_url(spreadsheet_url)
        # 選擇對應的工作表
        worksheet = spreadsheet.worksheet_by_title(sheet_title)
        # 讀取工作表內容並轉換為 Pandas DataFrame
        df = worksheet.get_as_df()
        # 獲取 DataFrame 的形狀
        num_rows = df.shape[0]
        # 新數據追加到下一行
        start_row = num_rows + 1

        data = [current_time, word, ','.join(pos['pos'] for pos in pos_list),
                ','.join(pos['chinese_definition'] for pos in pos_list),
                ','.join(example for example in example_list), us_pronunciation_url if us_pronunciation_url else '',
                uk_pronunciation_url if uk_pronunciation_url else '']

        # 在指定的行數插入新數據
        worksheet.insert_rows(start_row, values=[data])
        return start_row

    except Exception as e:
        # 如果儲存失敗，引發自定義的 SaveCardError 並帶上錯誤訊息
        raise SaveCardError("卡片儲存失敗，請稍後再試。")


# 函數：尋找使用者所有工作表
def get_user_worksheets(user_id, spreadsheet_urls, service_file_path):
    # 設定 Google Sheets API 的授權檔案（輸入金鑰）
    gc = pygsheets.authorize(service_file=service_file_path)
    # 建立一個字典用來存儲各個 spreadsheet 中對應的 user_worksheets
    user_worksheets_dict = {}

    # 逐一處理每個 spreadsheet
    for spreadsheet_url in spreadsheet_urls:
        spreadsheet = gc.open_by_url(spreadsheet_url)

        # 找到使用者對應的工作表
        user_worksheets = [worksheet.title.split('_')[1] for worksheet in spreadsheet.worksheets() if
                           user_id in worksheet.title]
        # 將結果存入字典
        user_worksheets_dict[spreadsheet.title] = user_worksheets
    return user_worksheets_dict

# 函數：反查卡片盒類型(工作表對照的資料庫)
def find_spreadsheet_by_worksheet(worksheet_name, spreadsheet_dict):
    for spreadsheet, worksheets in spreadsheet_dict.items():
        if worksheet_name in worksheets:
            return spreadsheet

# 函數：反查單字卡片盒裡面的內容
def process_flashcard_deck_v1(all_data, column_names):
    current_time_list, word_list, pos_list, chinese_list, example_list, note_list = ([] for _ in range(6))

    # 將數據分配到相應的列表中
    for row in all_data[1:]:
        if any(row):  # 檢查行是否包含有效數據
            current_time_list.append(row[column_names.index('新增時間')])
            word_list.append(row[column_names.index('單字')])
            pos_list.append(row[column_names.index('詞性')])
            chinese_list.append(row[column_names.index('中文')])
            example_list.append(row[column_names.index('例句')])
            note_list.append(row[column_names.index('筆記')])

    return current_time_list, word_list, pos_list, chinese_list, example_list, note_list

# 函數：一般查看單字卡
def generate_flex_message(current_time, word_name, pos_list, chinese_list, example_list, note_list):
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": "單字卡",
                    "weight": "bold",
                    "color": "#1DB446",
                    "size": "sm"
                },
                {
                    "type": "text",
                    "text": word_name,
                    "weight": "bold",
                    "size": "xxl",
                    "margin": "md"
                },
                {
                    "type": "text",
                    "text": pos_list,
                    "color": "#aaaaaa",
                    "size": "sm",
                    "margin": "sm"
                },
                {
                    "type": "separator",
                    "margin": "lg"
                },
                {
                    "type": "box",
                    "layout": "vertical",
                    "margin": "lg",
                    "spacing": "sm",
                    "contents": [
                        {
                            "type": "box",
                            "layout": "baseline",
                            "spacing": "sm",
                            "contents": [
                                {
                                    "type": "text",
                                    "text": "中文",
                                    "color": "#aaaaaa",
                                    "size": "sm",
                                    "flex": 1
                                },
                                {
                                    "type": "text",
                                    "text": chinese_list,
                                    "wrap": True,
                                    "color": "#666666",
                                    "size": "sm",
                                    "flex": 5
                                }
                            ]
                        },
                        {
                            "type": "box",
                            "layout": "baseline",
                            "spacing": "sm",
                            "contents": [
                                {
                                    "type": "text",
                                    "text": "例句",
                                    "color": "#aaaaaa",
                                    "size": "sm",
                                    "flex": 1
                                },
                                {
                                    "type": "text",
                                    "text": example_list,
                                    "wrap": True,
                                    "color": "#666666",
                                    "size": "sm",
                                    "flex": 5
                                }
                            ]
                        },
                        {
                            "type": "box",
                            "layout": "baseline",
                            "spacing": "sm",
                            "contents": [
                                {
                                    "type": "text",
                                    "text": "筆記",
                                    "color": "#aaaaaa",
                                    "size": "sm",
                                    "flex": 1
                                },
                                {
                                    "type": "text",
                                    "text": note_list,
                                    "wrap": True,
                                    "color": "#666666",
                                    "size": "sm",
                                    "flex": 5
                                }
                            ]
                        },
                        {
                            "type": "separator",
                            "margin": "xl"
                        },
                        {
                            "type": "box",
                            "layout": "baseline",
                            "spacing": "sm",
                            "contents": [
                                {
                                    "type": "text",
                                    "text": f'建立日期 {current_time}',
                                    "color": "#aaaaaa",
                                    "size": "xs",
                                    "flex": 2,
                                    "align": "end",
                                    "margin": "none"
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    }

def generate_see_more_bubble():
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "sm",
            "contents": [
                {
                    "type": "button",
                    "flex": 1,
                    "gravity": "center",
                    "action": {
                        "type": "uri",
                        "label": "See more",
                        "uri": "https://linecorp.com"
                    }
                }
            ]
        }
    }

# 使用者狀態、卡片盒字典、front.back_input
user_states = {}
user_decks = {}
user_word_decks = {}
user_searching_word_decks = {}
user_content = {}
user_insert_content = {}
user_new_word_content = {}
user_insert_word_content = {}
user_states_content = {}
user_searching_words = {}
user_decks_name = {}


# 處理訊息事件的函數
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    # 取得使用者資訊
    global data_lists
    user_id = event.source.user_id
    user_input = event.message.text.strip()

    """建立卡片功能"""
    # 判斷使用者輸入
    if '自建卡片' in user_input:
        text_message = TextSendMessage(text="""歡迎使用建立卡片功能！
請選擇您要建立的卡片類型：
👉單字卡：包含單字、詞性、中文解釋以及例句
👉閃卡：可以利用卡片的正面和反面輸入需要複習的內容""")
        bubble_message = FlexSendMessage(
            alt_text='選擇卡片類型',
            contents={
                "type": "bubble",
                "body": {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [
                        {
                            "type": "text",
                            "text": "歡迎使用建立卡片功能！",
                            "weight": "bold",
                            "size": "md"
                        },
                        {
                            "type": "text",
                            "text": "請選擇您要建立的卡片類型：",
                            "size": "sm",
                            "wrap": True
                        }
                    ]
                },
                "footer": {
                    "type": "box",
                    "layout": "horizontal",
                    "spacing": "sm",
                    "contents": [
                        {
                            "type": "button",
                            "style": "secondary",
                            "height": "sm",
                            "action": {
                                "type": "message",
                                "label": "單字卡",
                                "text": "我要建立單字卡"
                            }
                        },
                        {
                            "type": "button",
                            "style": "secondary",
                            "height": "sm",
                            "action": {
                                "type": "message",
                                "label": "閃卡",
                                "text": "我要建立閃卡"
                            }
                        }
                    ],
                    "flex": 0
                }
            }
        )
        # 回覆訊息
        line_bot_api.reply_message(event.reply_token, [text_message, bubble_message])
        user_states[user_id] = 'waiting_for_choosing_type'

    ######閃卡卡片盒######
    elif user_id in user_states and user_states[user_id] == 'waiting_for_choosing_type' and user_input == '我要建立閃卡':
        reply_text = '請輸入閃卡卡片盒名稱'
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_deck_name'
    # 尋找單字卡片盒名稱(使用現有卡片盒/建立新卡片盒）
    elif user_id in user_states and user_states[user_id] == 'waiting_for_deck_name':
        deck_name = user_input
        service_file_path = './client_secret.json'
        spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing'
        all_worksheets = get_all_worksheets(spreadsheet_url, service_file_path, user_id, deck_name)

        if all_worksheets is not None:
            # 提示使用者確認是否使用現有卡片盒
            reply_text = f'已有閃卡卡片盒「{deck_name}」'
            confirm_message = FlexSendMessage(
                alt_text='新增卡片至卡片盒確認',
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": f"已有此閃卡卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要使用閃卡卡片盒「{deck_name}」？",
                                "size": "sm",
                                "wrap": True
                            }
                        ]
                    },
                    "footer": {
                        "type": "box",
                        "layout": "horizontal",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "是",
                                    "text": "是"
                                }
                            },
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "否",
                                    "text": "否"
                                }
                            }
                        ]
                    }
                }
            )
            line_bot_api.reply_message(event.reply_token, [TextSendMessage(text=reply_text), confirm_message])
            user_decks[user_id] = deck_name  # 儲存使用者選擇的卡片盒
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_confirm_existing_deck'
        else:
            # 提示使用者確認是否建立新卡片盒
            reply_text = f'找不到閃卡卡片盒「{deck_name}」'
            confirm_message = FlexSendMessage(
                alt_text='新增卡片盒確認',
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": f"找不到閃卡卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要建立閃卡卡片盒「{deck_name}」？",
                                "size": "sm",
                                "wrap": True
                            }
                        ]
                    },
                    "footer": {
                        "type": "box",
                        "layout": "horizontal",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "是",
                                    "text": "Yes"
                                }
                            },
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "否",
                                    "text": "No"
                                }
                            }
                        ]
                    }
                }
            )
            line_bot_api.reply_message(event.reply_token, [TextSendMessage(text=reply_text), confirm_message])
            user_decks[user_id] = deck_name  # 儲存使用者選擇的卡片盒
            user_states.pop(user_id, None)  # 清除用戶狀態
            user_states[user_id] = 'waiting_for_confirm_new_deck'

    # 使用者確認是否使用現有卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_existing_deck':
        if user_input == '是':
            # 使用者確認使用現有卡片盒
            reply_text = f'確定使用閃卡卡片盒「{user_decks[user_id]}」'
            reply_text2 = """請輸入卡片內容：
（先輸入卡片正面，再換行輸入卡片背面）
———————————————
✍️格式範例：
什麼是卡片盒機器人？
Flashcards卡片盒機器人是一款幫助學習的Line帳號！"""

            # 創建兩條回覆訊息
            message1 = TextSendMessage(text=reply_text)
            message2 = TextSendMessage(text=reply_text2)

            # 一次性回覆兩條消息
            line_bot_api.reply_message(event.reply_token, [message1, message2])
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_user_input_content'
        elif user_input == "否":
            reply_text = f'取消使用閃卡卡片盒「{user_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 使用者輸入卡片內容(插入)
    elif user_id in user_states and user_states[user_id] == 'waiting_for_user_input_content':
        insert_content_list = user_input.split("\n")
        # 檢查使用者輸入格式
        if len(insert_content_list) != 2:
            # 格式不正確，提醒使用者
            reply_text = "格式輸入錯誤，請先輸入卡片正面，換行後再輸入卡片背面。"
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            return  # 結束處理，等待用戶重新輸入
        # 動態生成 Flex Message JSON
        flex_message = {
            "type": "bubble",
            "body": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "text",
                        "text": "建立閃卡",
                        "color": "#1DB446",
                        "size": "sm",
                        "weight": "bold"
                    },
                    {
                        "type": "text",
                        "text": user_decks[user_id],  # 卡片盒名稱
                        "weight": "bold",
                        "size": "xxl",
                        "margin": "md"
                    },
                    {
                        "type": "separator",
                        "margin": "xl"
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "margin": "lg",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "卡片正面",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 2
                                    },
                                    {
                                        "type": "text",
                                        "text": insert_content_list[0],  # 使用者輸入的第一行作為卡片正面
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            },
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "卡片背面",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 2
                                    },
                                    {
                                        "type": "text",
                                        "text": insert_content_list[1],  # 使用者輸入的第二行作為卡片背面
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "separator",
                        "margin": "xl"
                    }
                ]
            },
            "footer": {
                "type": "box",
                "layout": "vertical",
                "spacing": "sm",
                "contents": [
                    {
                        "type": "button",
                        "style": "primary",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "建立至卡片盒",
                            "text": "新增至閃卡卡片盒"
                        }
                    },
                    {
                        "type": "button",
                        "style": "link",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "取消",
                            "text": "取消新增至閃卡卡片盒"
                        }
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [],
                        "margin": "sm"
                    }
                ],
                "flex": 0
            }
        }

        # 使用 Flex Message 回覆使用者
        flex_reply_message = FlexSendMessage(alt_text="建立閃卡", contents=flex_message)
        line_bot_api.reply_message(event.reply_token, flex_reply_message)
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_insert_new_content'
        user_insert_content[user_id] = insert_content_list

    # 將卡片內容插入現有的卡片盒裡面
    elif user_id in user_states and user_states[user_id] == 'waiting_for_insert_new_content':
        if user_input == "新增至閃卡卡片盒":
            sheet_title = f'{user_id}_{user_decks[user_id]}'
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing'
            tz = pytz.timezone('Asia/Taipei')
            current_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            try:
                # 插入卡片內容並取得新插入資料的位置
                insert_position = insert_card_content_to_sheet(current_time, sheet_title, user_insert_content[user_id],
                                                               service_file_path, spreadsheet_url)
                # 提示使用者確認是否儲存卡片內容以及新插入資料的位置
                reply_text = f'已成功新增第 {insert_position} 張卡片至閃卡卡片盒「{user_decks[user_id]}」'
                message = TextSendMessage(text=reply_text)
                line_bot_api.reply_message(event.reply_token, message)
                user_states.pop(user_id, None)
            except SaveCardError as e:
                # 捕捉自定義的例外狀況並向使用者發送錯誤提示
                error_message = TextSendMessage(text=str(e))
                line_bot_api.reply_message(event.reply_token, error_message)
                user_states.pop(user_id, None)

        elif user_input == "取消新增至閃卡卡片盒":
            reply_text = f'取消建立至閃卡卡片盒「{user_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 使用者確認是否建立新閃卡卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_new_deck':
        if user_input == 'Yes':
            # 使用者確認建立新卡片盒
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing'
            create_new_worksheet(user_id, user_decks[user_id], service_file_path, spreadsheet_url)

            # 新工作表建立成功的提示
            reply_text = f'確定建立閃卡卡片盒「{user_decks[user_id]}」'
            reply_text2 = """請輸入卡片內容：
（先輸入卡片正面，再換行輸入卡片背面）
———————————————
✍️格式範例：
什麼是卡片盒機器人？
Flashcards卡片盒機器人是一款幫助學習的Line帳號！"""

            # 創建兩條回覆訊息
            message1 = TextSendMessage(text=reply_text)
            message2 = TextSendMessage(text=reply_text2)

            # 一次性回覆兩條消息
            line_bot_api.reply_message(event.reply_token, [message1, message2])
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_user_input_new_content'

        elif user_id == "No":
            reply_text = f'取消建立閃卡卡片盒「{user_decks[user_id]}」'
            # 回覆使用者
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 使用者輸入單字卡內容(新)
    elif user_id in user_states and user_states[user_id] == 'waiting_for_user_input_new_content':
        content_list = user_input.split("\n")
        # 檢查使用者輸入格式
        if len(content_list) != 2:
            # 格式不正確，提醒使用者
            reply_text = "格式輸入錯誤，請先輸入卡片正面，換行後再輸入卡片背面。"
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            return  # 結束處理，等待用戶重新輸入

        # 動態生成 Flex Message JSON
        flex_message = {
            "type": "bubble",
            "body": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "text",
                        "text": "建立閃卡",
                        "color": "#1DB446",
                        "size": "sm",
                        "weight": "bold"
                    },
                    {
                        "type": "text",
                        "text": user_decks[user_id],  # 卡片盒名稱
                        "weight": "bold",
                        "size": "xxl",
                        "margin": "md"
                    },
                    {
                        "type": "separator",
                        "margin": "xl"
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "margin": "lg",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "卡片正面",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 2
                                    },
                                    {
                                        "type": "text",
                                        "text": content_list[0],  # 使用者輸入的第一行作為卡片正面
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            },
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "卡片背面",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 2
                                    },
                                    {
                                        "type": "text",
                                        "text": content_list[1],  # 使用者輸入的第二行作為卡片背面
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "separator",
                        "margin": "xl"
                    }
                ]
            },
            "footer": {
                "type": "box",
                "layout": "vertical",
                "spacing": "sm",
                "contents": [
                    {
                        "type": "button",
                        "style": "primary",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "建立至卡片盒",
                            "text": "建立至閃卡卡片盒"
                        }
                    },
                    {
                        "type": "button",
                        "style": "link",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "取消",
                            "text": "取消建立至閃卡卡片盒"
                        }
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [],
                        "margin": "sm"
                    }
                ],
                "flex": 0
            }
        }

        # 使用 Flex Message 回覆使用者
        flex_reply_message = FlexSendMessage(alt_text="建立閃卡", contents=flex_message)
        line_bot_api.reply_message(event.reply_token, flex_reply_message)
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_save_new_content'
        user_content[user_id] = content_list

    # 將卡片內容插入新的卡片盒裡面
    elif user_id in user_states and user_states[user_id] == 'waiting_for_save_new_content':
        if user_input == "建立至閃卡卡片盒":
            sheet_title = f'{user_id}_{user_decks[user_id]}'
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing'
            tz = pytz.timezone('Asia/Taipei')
            current_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            try:
                result = save_card_content_to_sheet(current_time, sheet_title, user_content[user_id], service_file_path,
                                                    spreadsheet_url)
                if result:
                    reply_text = f'已成功儲存卡片至閃卡卡片盒「{user_decks[user_id]}」'
                    message = TextSendMessage(text=reply_text)
                    line_bot_api.reply_message(event.reply_token, message)
                    user_states.pop(user_id, None)

            except SaveCardError as e:
                # 捕捉自定義的例外狀況並向使用者發送錯誤提示
                error_message = TextSendMessage(text=str(e))
                line_bot_api.reply_message(event.reply_token, error_message)
                user_states.pop(user_id, None)

        elif user_input == "取消建立至閃卡卡片盒":
            reply_text = f'取消建立至閃卡卡片盒「{user_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    ######單字卡卡片盒#######
    elif user_id in user_states and user_states[
        user_id] == 'waiting_for_choosing_type' and user_input == '我要建立單字卡':
        reply_text = '請輸入單字卡卡片盒名稱'
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_word_deck_name'
    # 尋找單字卡片盒名稱(使用現有卡片盒/建立新卡片盒）
    elif user_id in user_states and user_states[user_id] == 'waiting_for_word_deck_name':
        word_deck_name = user_input
        service_file_path = './client_secret.json'
        spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1_0JteKeNM4yf3QUMKc8R3qpCQTgEhq7K7jfPHnlizio/edit?usp=sharing'
        all_worksheets = get_all_worksheets(spreadsheet_url, service_file_path, user_id, word_deck_name)

        if all_worksheets is not None:
            # 提示使用者確認是否使用現有卡片盒
            reply_text = f'已有此單字卡卡片盒「{word_deck_name}」'
            confirm_message = FlexSendMessage(
                alt_text='新增卡片至單字卡卡片盒確認',
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": f"已有此單字卡卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要使用單字卡卡片盒「{word_deck_name}」？",
                                "size": "sm",
                                "wrap": True
                            }
                        ]
                    },
                    "footer": {
                        "type": "box",
                        "layout": "horizontal",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "是",
                                    "text": "確定"
                                }
                            },
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "否",
                                    "text": "不確定"
                                }
                            }
                        ]
                    }
                }
            )
            line_bot_api.reply_message(event.reply_token, [TextSendMessage(text=reply_text), confirm_message])
            user_word_decks[user_id] = word_deck_name  # 儲存使用者輸入的卡片盒名稱
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_confirm_existing_word_deck'
        else:
            # 提示使用者確認是否建立新單字卡片盒
            reply_text = f'找不到單字卡卡片盒「{word_deck_name}」'
            confirm_message = FlexSendMessage(
                alt_text='新增單字卡卡片盒確認',
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": f"找不到單字卡卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要建立單字卡卡片盒「{word_deck_name}」？",
                                "size": "sm",
                                "wrap": True
                            }
                        ]
                    },
                    "footer": {
                        "type": "box",
                        "layout": "horizontal",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "是",
                                    "text": "好"
                                }
                            },
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "否",
                                    "text": "不好"
                                }
                            }
                        ]
                    }
                }
            )
            line_bot_api.reply_message(event.reply_token, [TextSendMessage(text=reply_text), confirm_message])
            user_word_decks[user_id] = word_deck_name  # 儲存使用者選擇的卡片盒
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_confirm_new_word_deck'

    # 使用者確認是否使用現有單字卡卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_existing_word_deck':
        if user_input == '確定':
            reply_text = f'確定使用單字卡卡片盒「{user_word_decks[user_id]}」'
            reply_text2 = """請輸入卡片內容：
(依序換行輸入單字、詞性、中文、例句及筆記，若無則填寫無)
———————————————
✍️格式範例：
flashcard
noun
閃卡/抽認卡
無
flashcard/flash card"""

            # 兩條回覆訊息
            message1 = TextSendMessage(text=reply_text)
            message2 = TextSendMessage(text=reply_text2)

            # 一次性回覆兩條消息
            line_bot_api.reply_message(event.reply_token, [message1, message2])
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_user_input_word_content'
        elif user_input == '不確定':
            reply_text = f'取消使用單字卡卡片盒「{user_word_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 使用者輸入單字卡內容(插入)
    elif user_id in user_states and user_states[user_id] == 'waiting_for_user_input_word_content':
        insert_word_content_list = user_input.split("\n")
        # 檢查使用者輸入格式
        if len(insert_word_content_list) != 5:
            reply_text = "格式輸入錯誤，請依序換行輸入單字、詞性、中文、例句以及筆記，若無請填寫無"
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            return  # 結束處理，等待用戶重新輸入

        # 動態生成 Flex Message JSON
        flex_message = {
            "type": "bubble",
            "body": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "text",
                        "text": "建立單字卡",
                        "color": "#1DB446",
                        "weight": "bold",
                        "size": "sm"
                    },
                    {
                        "type": "text",
                        "weight": "bold",
                        "size": "xxl",
                        "margin": "md",
                        "text": insert_word_content_list[0]  # 單字
                    },
                    {
                        "type": "text",
                        "text": insert_word_content_list[1],  # 詞性
                        "color": "#aaaaaa",
                        "size": "sm",
                        "margin": "sm"
                    },
                    {
                        "type": "separator",
                        "margin": "lg"
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "margin": "lg",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "中文",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 1
                                    },
                                    {
                                        "type": "text",
                                        "text": insert_word_content_list[2],  # 中文
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            },
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "例句",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 1
                                    },
                                    {
                                        "type": "text",
                                        "text": insert_word_content_list[3],  # 例句
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            },
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "筆記",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 1
                                    },
                                    {
                                        "type": "text",
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5,
                                        "text": insert_word_content_list[4]  # 筆記
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "separator",
                        "margin": "xl"
                    }
                ]
            },
            "footer": {
                "type": "box",
                "layout": "vertical",
                "spacing": "sm",
                "contents": [
                    {
                        "type": "button",
                        "style": "primary",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "建立至卡片盒",
                            "text": "建立至單字卡卡片盒"
                        }
                    },
                    {
                        "type": "button",
                        "style": "link",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "取消",
                            "text": "取消建立至單字卡卡片盒"
                        }
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [],
                        "margin": "sm"
                    }
                ],
                "flex": 0
            }
        }

        # 使用 Flex Message 回覆使用者
        flex_reply_message = FlexSendMessage(alt_text="建立單字卡", contents=flex_message)
        line_bot_api.reply_message(event.reply_token, flex_reply_message)
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_insert_new_word_content'
        user_insert_word_content[user_id] = insert_word_content_list

    # 將卡片內容插入現有的卡片盒裡面
    elif user_id in user_states and user_states[user_id] == 'waiting_for_insert_new_word_content':
        if user_input == "建立至單字卡卡片盒":
            sheet_title = f'{user_id}_{user_word_decks[user_id]}'
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1_0JteKeNM4yf3QUMKc8R3qpCQTgEhq7K7jfPHnlizio/edit?usp=sharing'
            tz = pytz.timezone('Asia/Taipei')
            current_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            try:
                # 插入卡片內容並取得新插入資料的位置
                insert_position = insert_word_card_content_to_sheet(current_time, sheet_title,
                                                                    user_insert_word_content[user_id],
                                                                    service_file_path, spreadsheet_url)
                # 提示使用者確認是否儲存卡片內容以及新插入資料的位置
                reply_text = f'已成功新增第 {insert_position} 個單字至單字卡卡片盒「{user_word_decks[user_id]}」'
                message = TextSendMessage(text=reply_text)
                line_bot_api.reply_message(event.reply_token, message)
                user_states.pop(user_id, None)
            except SaveCardError as e:
                # 捕捉自定義的例外狀況並向使用者發送錯誤提示
                error_message = TextSendMessage(text=str(e))
                line_bot_api.reply_message(event.reply_token, error_message)
                user_states.pop(user_id, None)

        elif user_input == "取消建立至單字卡卡片盒":
            reply_text = f'取消建立至單字卡卡片盒「{user_word_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 使用者確認是否建立新單字卡卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_new_word_deck':
        if user_input == '好':
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1_0JteKeNM4yf3QUMKc8R3qpCQTgEhq7K7jfPHnlizio/edit?usp=sharing'
            create_new_worksheet(user_id, user_word_decks[user_id], service_file_path, spreadsheet_url)

            # 更新使用者狀態
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_user_input_new_word_content'

            # 新工作表建立成功的提示
            reply_text = f'確定建立單字卡卡片盒「{user_word_decks[user_id]}」'
            reply_text2 = """請輸入卡片內容：
(依序換行輸入單字、詞性、中文、例句及筆記，若無則填寫無)
———————————————
✍️格式範例：
flashcard
noun
閃卡/抽認卡
無
flashcard/flash card"""

            # 創建兩條回覆訊息
            message1 = TextSendMessage(text=reply_text)
            message2 = TextSendMessage(text=reply_text2)

            # 一次性回覆兩條消息
            line_bot_api.reply_message(event.reply_token, [message1, message2])

        elif user_input == '不好':
            reply_text = f'取消建立單字卡卡片盒「{user_word_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 使用者輸入單字卡內容(新)
    elif user_id in user_states and user_states[user_id] == 'waiting_for_user_input_new_word_content':
        new_word_content_list = user_input.split("\n")
        # 檢查使用者輸入格式
        if len(new_word_content_list) != 5:
            # 格式不正確，提醒使用者
            reply_text = "格式輸入錯誤，請依序換行輸入單字、詞性、中文、例句以及筆記，若無請填寫無"
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            return  # 結束處理，等待用戶重新輸入

        # 動態生成 Flex Message JSON
        flex_message = {
            "type": "bubble",
            "body": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "text",
                        "text": "建立單字卡",
                        "color": "#1DB446",
                        "weight": "bold",
                        "size": "sm"
                    },
                    {
                        "type": "text",
                        "weight": "bold",
                        "size": "xxl",
                        "margin": "md",
                        "text": new_word_content_list[0]  # 單字
                    },
                    {
                        "type": "text",
                        "text": new_word_content_list[1],  # 詞性
                        "color": "#aaaaaa",
                        "size": "sm",
                        "margin": "sm"
                    },
                    {
                        "type": "separator",
                        "margin": "lg"
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "margin": "lg",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "中文",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 1
                                    },
                                    {
                                        "type": "text",
                                        "text": new_word_content_list[2],  # 中文
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            },
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "例句",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 1
                                    },
                                    {
                                        "type": "text",
                                        "text": new_word_content_list[3],  # 例句
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5
                                    }
                                ]
                            },
                            {
                                "type": "box",
                                "layout": "baseline",
                                "spacing": "sm",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": "筆記",
                                        "color": "#aaaaaa",
                                        "size": "sm",
                                        "flex": 1
                                    },
                                    {
                                        "type": "text",
                                        "wrap": True,
                                        "color": "#666666",
                                        "size": "sm",
                                        "flex": 5,
                                        "text": new_word_content_list[4]  # 筆記
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "separator",
                        "margin": "xl"
                    }
                ]
            },
            "footer": {
                "type": "box",
                "layout": "vertical",
                "spacing": "sm",
                "contents": [
                    {
                        "type": "button",
                        "style": "primary",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "建立至卡片盒",
                            "text": "新增至單字卡卡片盒"
                        }
                    },
                    {
                        "type": "button",
                        "style": "link",
                        "height": "sm",
                        "action": {
                            "type": "message",
                            "label": "取消",
                            "text": "取消新增至單字卡卡卡片盒"
                        }
                    },
                    {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [],
                        "margin": "sm"
                    }
                ],
                "flex": 0
            }
        }

        # 使用 Flex Message 回覆使用者
        flex_reply_message = FlexSendMessage(alt_text="建立單字卡", contents=flex_message)
        line_bot_api.reply_message(event.reply_token, flex_reply_message)
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_save_new_word_content'
        user_new_word_content[user_id] = new_word_content_list

    # 將卡片內容插入新的單字卡片盒裡面
    elif user_id in user_states and user_states[user_id] == 'waiting_for_save_new_word_content':
        if user_input == "新增至單字卡卡片盒":
            sheet_title = f'{user_id}_{user_word_decks[user_id]}'
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1_0JteKeNM4yf3QUMKc8R3qpCQTgEhq7K7jfPHnlizio/edit?usp=sharing'
            tz = pytz.timezone('Asia/Taipei')
            current_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            try:
                result = save_word_card_content_to_sheet(current_time, sheet_title, user_new_word_content[user_id],
                                                         service_file_path,
                                                         spreadsheet_url)
                if result:
                    reply_text = f'已成功儲存卡片至單字卡卡片盒「{user_word_decks[user_id]}」'
                    message = TextSendMessage(text=reply_text)
                    line_bot_api.reply_message(event.reply_token, message)
                    user_states.pop(user_id, None)
            except SaveCardError as e:
                # 捕捉自定義的例外狀況並向使用者發送錯誤提示
                error_message = TextSendMessage(text=str(e))
                line_bot_api.reply_message(event.reply_token, error_message)
                user_states.pop(user_id, None)

        elif user_input == "取消新增至單字卡卡片盒":
            reply_text = f'取消新增至單字卡卡片盒「{user_word_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    """查單字功能"""
    if '查單字' in user_input:
        reply_text = '請輸入想要查詢的英文單字'
        # 回覆使用者
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)
        user_states[user_id] = 'waiting_for_searching_word'
    # 查字典功能
    elif user_id in user_states and user_states[user_id] == 'waiting_for_searching_word':
        searching_word = user_input
        # 爬蟲回傳的字典串列，pos_list是flexmessage需要的詞性及解釋，example_list是replymessage需要的例句
        pos_list, example_list, us_pron_url, uk_pron_url = lookup_word(searching_word)
        if not pos_list:
            # 如果查詢不到結果，回覆相應訊息給使用者
            reply_text = f'抱歉，找不到單字 "{searching_word}" 的相關資訊，請嘗試輸入其他單字。'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)
        else:
            flex_contents = []  # 整理flex message的變動資料(依據各個單字詞性多寡跑迴圈)
            for i in pos_list:
                part_of_speech = i["pos"]
                definition = i["chinese_definition"]
                obj = {
                    "type": "box",
                    "layout": "horizontal",
                    "contents": [
                        {
                            "type": "text",
                            "text": part_of_speech,
                            "size": "sm",
                            "color": "#555555",
                            "flex": 0,
                        },
                        {
                            "type": "text",
                            "text": definition,
                            "wrap": True,
                            "size": "sm",
                            "color": "#111111",
                            "align": "end",
                        },
                    ],
                }

                flex_contents.append(obj)

            # 整理flex message固定內容的資料
            fixed_contents = [
                {
                    "type": "box",
                    "layout": "horizontal",
                    "contents": [
                        {
                            "type": "text",
                            "text": "參考字典",
                            "size": "sm",
                            "color": "#555555",
                            "flex": 0,
                        },
                        {
                            "type": "text",
                            "text": "劍橋字典",
                            "size": "sm",
                            "color": "#111111",
                            "align": "end",
                        },
                    ],
                },
                {"type": "separator", "margin": "xxl"},
                {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [
                        {
                            "type": "button",
                            "action": {
                                "type": "uri",
                                "label": "聽美式發音",
                                "uri": us_pron_url,  # 使用劍橋字典的美式發音連結
                            },
                        },
                        {
                            "type": "button",
                            "action": {
                                "type": "uri",
                                "label": "聽英式發音",
                                "uri": uk_pron_url,  # 使用劍橋字典的英式發音連結
                            },
                        },
                        {
                            "type": "button",
                            "action": {
                                "type": "message",
                                "label": "查看例句",
                                "text": f"查看例句 {searching_word}",
                            },
                        },
                        {
                            "type": "button",
                            "action": {
                                "type": "message",
                                "label": "建立卡片",
                                "text": "建立字卡",
                            },
                        },
                    ],
                },
            ]
            flex_contents.extend(fixed_contents)

            # ... (其餘程式碼)

            flex_message = FlexSendMessage(
                alt_text=user_input,
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": "查字典",
                                "weight": "bold",
                                "color": "#1DB446",
                                "size": "sm",
                            },
                            {
                                "type": "text",
                                "text": searching_word,
                                "weight": "bold",
                                "size": "xxl",
                                "margin": "md",
                            },
                            {"type": "separator", "margin": "xxl"},
                            {
                                "type": "box",
                                "layout": "vertical",
                                "margin": "xxl",
                                "spacing": "sm",
                                "contents": flex_contents,
                            },
                        ],
                    },
                    "styles": {"footer": {"separator": True}},
                },
            )
            line_bot_api.reply_message(event.reply_token, flex_message)
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_choosing_button'
            user_searching_words[user_id] = searching_word

    # 字典產生後，使用者選擇按鈕（查看例句/建立字卡）
    elif user_id in user_states and user_states[user_id] == 'waiting_for_choosing_button':
        if "查看例句" in user_input:
            pos_list, example_list, us_pron_url, uk_pron_url = lookup_word(user_searching_words[user_id])
            send_message_list = []  # Linebot要一次發送多個訊息需要先把訊息用list包起來
            for reply_example in example_list:
                if len(send_message_list) < 5:  # Linebot一次發送訊息不能超過五則
                    send_message_list.append(
                        TextSendMessage(text=f"{user_searching_words[user_id]}\n{reply_example}")
                    )
            line_bot_api.reply_message(event.reply_token, send_message_list)

        if "建立字卡" in user_input:
            reply_text = '請輸入字典卡片盒名稱'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_searching_word_deck_name'

    # 尋找字典卡片盒名稱（使用現有名稱/建立新名稱）
    elif user_id in user_states and user_states[user_id] == 'waiting_for_searching_word_deck_name':
        searching_word_deck_name = user_input
        service_file_path = './client_secret.json'
        spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1UdUmyvZ-W1kkIohnoRxHEd5ISF9PkbF-RtBxtdXteaU/edit?usp=sharing'
        all_worksheets = get_all_worksheets(spreadsheet_url, service_file_path, user_id, searching_word_deck_name)

        if all_worksheets is not None:
            # 提示使用者確認是否使用現有卡片盒
            reply_text = f'已有此字典卡片盒「{searching_word_deck_name}」'
            confirm_message = FlexSendMessage(
                alt_text='新增卡片至字典卡片盒確認',
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": f"已有此字典卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要使用字典卡片盒「{searching_word_deck_name}」？",
                                "size": "sm",
                                "wrap": True
                            }
                        ]
                    },
                    "footer": {
                        "type": "box",
                        "layout": "horizontal",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "是",
                                    "text": "要"
                                }
                            },
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "否",
                                    "text": "不要"
                                }
                            }
                        ]
                    }
                }
            )
            line_bot_api.reply_message(event.reply_token, [TextSendMessage(text=reply_text), confirm_message])
            user_searching_word_decks[user_id] = searching_word_deck_name  # 儲存使用者選擇的卡片盒
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_confirm_existing_searching_word_deck'
        else:
            # 提示使用者確認是否建立新卡片盒
            reply_text = f'找不到字典卡片盒「{searching_word_deck_name}」'
            confirm_message = FlexSendMessage(
                alt_text='新增字典卡片盒確認',
                contents={
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": f"找不到字典卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要建立字典卡片盒「{searching_word_deck_name}」？",
                                "size": "sm",
                                "wrap": True
                            }
                        ]
                    },
                    "footer": {
                        "type": "box",
                        "layout": "horizontal",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "是",
                                    "text": "Y"
                                }
                            },
                            {
                                "type": "button",
                                "style": "secondary",
                                "height": "sm",
                                "action": {
                                    "type": "message",
                                    "label": "否",
                                    "text": "N"
                                }
                            }
                        ]
                    }
                }
            )
            line_bot_api.reply_message(event.reply_token, [TextSendMessage(text=reply_text), confirm_message])
            user_searching_word_decks[user_id] = searching_word_deck_name
            user_states.pop(user_id, None)
            user_states[user_id] = 'waiting_for_confirm_new_searching_word_deck'

    # 將字典加入新的字典卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_new_searching_word_deck':
        if user_input == 'Y':
            reply_text1 = f'確定建立字典卡片盒「{user_searching_word_decks[user_id]}」'
            sheet_title = f'{user_id}_{user_searching_word_decks[user_id]}'
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1UdUmyvZ-W1kkIohnoRxHEd5ISF9PkbF-RtBxtdXteaU/edit?usp=sharing'
            tz = pytz.timezone('Asia/Taipei')
            current_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            try:
                pos_list, example_list, us_pronunciation_url, uk_pronunciation_url = lookup_word(
                    user_searching_word_decks[user_id])
                result = searching_word_to_sheet(current_time, service_file_path, spreadsheet_url, sheet_title,
                                                 user_searching_word_decks[user_id], pos_list, example_list,
                                                 us_pronunciation_url, uk_pronunciation_url)
                # 提示使用者卡片是否成功儲存
                if result:
                    reply_text2 = f'已成功儲存卡片至字典卡片盒「{user_searching_word_decks[user_id]}」'
                    message1 = TextSendMessage(text=reply_text1)
                    message2 = TextSendMessage(text=reply_text2)
                    line_bot_api.reply_message(event.reply_token, [message1, message2])
                    user_states.pop(user_id, None)

            except SaveCardError as e:
                # 捕捉自定義的例外狀況並向使用者發送錯誤提示
                error_message = TextSendMessage(text=str(e))
                line_bot_api.reply_message(event.reply_token, error_message)
                user_states.pop(user_id, None)
        elif user_input == 'N':
            reply_text = f'取消建立字典卡片盒「{user_searching_word_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    # 將字典加入現有字典卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_existing_searching_word_deck':
        if user_input == '要':
            reply_text1 = f'確定使用字典卡片盒「{user_searching_word_decks[user_id]}」'
            sheet_title = f'{user_id}_{user_searching_word_decks[user_id]}'
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1UdUmyvZ-W1kkIohnoRxHEd5ISF9PkbF-RtBxtdXteaU/edit?usp=sharing'
            tz = pytz.timezone('Asia/Taipei')
            current_time = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
            try:
                pos_list, example_list, us_pronunciation_url, uk_pronunciation_url = lookup_word(
                    user_searching_word_decks[user_id])
                insert_position = searching_word_to_existing_sheet(current_time, service_file_path, spreadsheet_url,
                                                                   sheet_title,
                                                                   user_searching_word_decks[user_id], pos_list,
                                                                   example_list,
                                                                   us_pronunciation_url, uk_pronunciation_url)
                # 提示使用者卡片是否成功儲存
                if insert_position:
                    reply_text2 = f'已成功新增第 {insert_position} 個單字至字典卡片盒「{user_searching_word_decks[user_id]}」'
                    message1 = TextSendMessage(text=reply_text1)
                    message2 = TextSendMessage(text=reply_text2)
                    line_bot_api.reply_message(event.reply_token, [message1, message2])
                    user_states.pop(user_id, None)

            except SaveCardError as e:
                # 捕捉自定義的例外狀況並向使用者發送錯誤提示
                error_message = TextSendMessage(text=str(e))
                line_bot_api.reply_message(event.reply_token, error_message)
                user_states.pop(user_id, None)
        elif user_input == "不要":
            reply_text = f'取消使用字典卡片盒「{user_searching_word_decks[user_id]}」'
            message = TextSendMessage(text=reply_text)
            line_bot_api.reply_message(event.reply_token, message)
            user_states.pop(user_id, None)

    """卡片盒功能"""
    if user_input == "卡片盒":
        reply_text = '請選擇卡片盒'
        spreadsheet_urls = [
            'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing',
            'https://docs.google.com/spreadsheets/d/1_0JteKeNM4yf3QUMKc8R3qpCQTgEhq7K7jfPHnlizio/edit?usp=sharing',
            'https://docs.google.com/spreadsheets/d/1UdUmyvZ-W1kkIohnoRxHEd5ISF9PkbF-RtBxtdXteaU/edit?usp=sharing'
        ]
        service_file_path = './client_secret.json'
        result = get_user_worksheets(user_id, spreadsheet_urls, service_file_path)
        # 儲存函數結果
        sheets_list = []
        lists_list = []
        for key, value in result.items():
            sheet = key
            sheets_list.append(sheet)
            lists_list.append(value)

        flex_contents = []  # 整理flex message的變動資料(依據各個資料庫的工作表內容來決定button數量及名稱)
        # 生成單字卡卡片盒的按鈕
        for i in lists_list[1]:
            flashcard_sheet_name = i
            obj = {
                "type": "button",
                "action": {
                    "type": "message",
                    "label": flashcard_sheet_name,
                    "text": f'單字卡卡片盒「{flashcard_sheet_name}」'
                }
            }
            flex_contents.append(obj)

        # 生成閃卡卡片盒的按鈕
        for i in lists_list[0]:
            word_sheet_name = i
            obj2 = {
                "type": "button",
                "action": {
                    "type": "message",
                    "label": word_sheet_name,
                    "text": f'閃卡卡片盒「{word_sheet_name}」'
                }
            }
            flex_contents.append(obj2)

        # 生成字典卡片盒的按鈕
        for i in lists_list[2]:
            dic_sheet_name = i
            obj3 = {
                "type": "button",
                "action": {
                    "type": "message",
                    "label": dic_sheet_name,
                    "text": f'字典卡片盒「{dic_sheet_name}」'
                }
            }
            flex_contents.append(obj3)

        # 計算每個部分的開始索引和結束索引(利用變數儲存比較清楚位置)
        start_flashcard = 0
        end_flashcard = len(lists_list[1])
        start_word = end_flashcard
        end_word = start_word + len(lists_list[0])
        start_dic = end_word

        # 使用計算出的索引來切割 flex_contents 列表
        flashcard_contents = flex_contents[start_flashcard:end_flashcard]
        word_contents = flex_contents[start_word:end_word]
        dic_contents = flex_contents[start_dic:]

        # 加到 bubble 的固定資料中
        carousel_content = {
            "type": "carousel",
            "contents": [
                {
                    "type": "bubble",
                    "hero": {
                        "type": "image",
                        "url": "https://img.onl/tZugX3",
                        "aspectRatio": "20:14",
                        "size": "full",
                        "aspectMode": "cover"
                    },
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": "✨單字卡卡片盒✨",
                                "weight": "bold",
                                "size": "lg",
                                "align": "center"
                            },
                            {
                                "type": "separator",
                                "margin": "xl"
                            },
                            *flashcard_contents
                        ]
                    }
                },
                {
                    "type": "bubble",
                    "hero": {
                        "type": "image",
                        "url": "https://img.onl/xdzcC",
                        "aspectRatio": "20:14",
                        "size": "full",
                        "aspectMode": "cover"
                    },
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": "✨閃卡卡片盒✨",
                                "weight": "bold",
                                "size": "lg",
                                "align": "center"
                            },
                            {
                                "type": "separator",
                                "margin": "xl"
                            },
                            *word_contents
                        ]
                    }
                },
                {
                    "type": "bubble",
                    "hero": {
                        "type": "image",
                        "url": "https://img.onl/Ir0ukQ",
                        "aspectRatio": "20:14",
                        "size": "full",
                        "aspectMode": "cover"
                    },
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                            {
                                "type": "text",
                                "text": "✨字典卡片盒✨",
                                "weight": "bold",
                                "size": "lg",
                                "align": "center"
                            },
                            {
                                "type": "separator",
                                "margin": "xl"
                            },
                            *dic_contents
                        ]
                    }
                }
            ]
        }

        message = TextSendMessage(text=reply_text)
        flex_message = FlexSendMessage(alt_text='Carousel Message', contents=carousel_content)
        line_bot_api.reply_message(event.reply_token, [message, flex_message])
        user_states[user_id] = 'choosing_card_box_mode'
    # 選擇carousel message上面的牌組，請使用者選擇學習模式
    elif user_id in user_states and user_states[user_id] == 'choosing_card_box_mode':
        card_box_name = user_input
        reply_text = f'已選擇{card_box_name}'
        bubble_message = FlexSendMessage(
            alt_text='選擇卡片類型',
            contents={
                "type": "bubble",
                "body": {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [
                        {
                            "type": "text",
                            "text": f'{card_box_name}',
                            "weight": "bold",
                            "size": "md"
                        },
                        {
                            "type": "text",
                            "text": f'請選擇您要學習的模式：',
                            "size": "sm",
                            "wrap": True
                        }
                    ]
                },
                "footer": {
                    "type": "box",
                    "layout": "horizontal",
                    "spacing": "sm",
                    "contents": [
                        {
                            "type": "button",
                            "style": "secondary",
                            "height": "sm",
                            "action": {
                                "type": "message",
                                "label": "一般查看",
                                "text": "查看卡片"
                            }
                        },
                        {
                            "type": "button",
                            "style": "secondary",
                            "height": "sm",
                            "action": {
                                "type": "message",
                                "label": "複習模式",
                                "text": "複習卡片"
                            }
                        }
                    ],
                    "flex": 0
                }
            }
        )
        # 回覆訊息
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, [message, bubble_message])
        user_states.pop(user_id, None)
        user_states[user_id] = 'waiting_for_choosing_mode'
        user_decks_name[user_id] = card_box_name

    # 選擇學習模式__一般查看
    elif user_id in user_states and user_states[user_id] == 'waiting_for_choosing_mode' and user_input == "查看卡片":
        deck_name = user_decks_name[user_id].split('「')[1].split('」')[0]
        # 提取「」以前的部分
        sheet_type = user_decks_name[user_id].split('「')[0]
        # 定義卡片盒名稱跟試算表URL
        sheet_type_mapping = {
            "單字卡卡片盒": 'https://docs.google.com/spreadsheets/d/1_0JteKeNM4yf3QUMKc8R3qpCQTgEhq7K7jfPHnlizio/edit?usp=sharing',
            "閃卡卡片盒": 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing',
            "字典卡片盒": 'https://docs.google.com/spreadsheets/d/1UdUmyvZ-W1kkIohnoRxHEd5ISF9PkbF-RtBxtdXteaU/edit?usp=sharing'
        }
        # 根據 sheet_type 獲取對應的試算表 URL
        sheet_url = sheet_type_mapping.get(sheet_type)
        sheet_name = f'{user_id}_{deck_name}'

        if sheet_url:
            # 初始化 spreadsheet
            gc = pygsheets.authorize(service_file='./client_secret.json')
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.worksheet_by_title(sheet_name)

            # 獲取所有數據
            all_data = worksheet.get_all_values()
            # 假設第一行是列名
            column_names = all_data[0]

            # 調用函數獲取數據
            current_time_list, word_list, pos_list, chinese_list, example_list, note_list = process_flashcard_deck_v1(
                all_data,
                column_names)

            columns_list = []
            data_lists = []
            # 將數據分開
            for name, data_list in zip(
                    ["Current Time List", "Word List", "POS List", "Chinese List", "Example List", "Note List"],
                    [current_time_list, word_list, pos_list, chinese_list, example_list, note_list]):
                columns_list.append(name)
                data_lists.append(data_list)

            # 只有在 data_lists 長度為 6 時，才生成 Flex Message
            if len(data_lists) == 6:
                # 動態生成 Flex Message JSON
                flex_messages = [generate_flex_message(current_time, word_name, pos_list, chinese_list, example_list,
                                                       note_list) for current_time, word_name, pos_list, chinese_list,
                                 example_list, note_list in
                                 zip(data_lists[0], data_lists[1], data_lists[2], data_lists[3],
                                     data_lists[4], data_lists[5])]

            if len(flex_messages) <= 10:
                # 少於等於 10 條 Bubble Messages，使用 Carousel Flex Message
                carousel_flex_message = FlexSendMessage(
                    alt_text="Carousel Flex Message",
                    contents={
                        "type": "carousel",
                        "contents": flex_messages
                    }
                )
            else:
                # 多於 10 條 Bubble Messages，使用 Carousel Flex Message 加上 See More 按鈕
                carousel_flex_message = FlexSendMessage(
                    alt_text="Carousel Flex Message",
                    contents={
                        "type": "carousel",
                        "contents": flex_messages[:9] + [generate_see_more_bubble()]
                    }
                )
            line_bot_api.reply_message(event.reply_token, carousel_flex_message)
            user_states.pop(user_id, None)


    else:
        # 其他操作失敗的情況
        reply_text = '機器人🤖讀取失敗，請重新嘗試'
        # 回覆使用者
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)


# 主程式
import os

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
