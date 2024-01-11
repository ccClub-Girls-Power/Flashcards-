#載入LineBot所需要的套件
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
line_bot_api = LineBotApi('Yxy/j1WMkLcW+qV2OJb1YIjx9V4sGCcUmxlP8FTtHiEkr/P5TiaOwocFqWnpeS1G1D7S/QO/Qibsd6P3u2pf7An14bzy/nb10NNfFtKgyIITsFNIB4Wl9o4xNzHk7Cgk+hab356oMaMA4gkt7tu26wdB04t89/1O/w1cDnyilFU=')
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


#訊息傳遞區塊
##### 基本上程式編輯都在這個function #####
# 載入必要的套件
from linebot.models import FlexSendMessage, TextSendMessage
import pygsheets
from linebot.models import MessageEvent, TextMessage

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

# 使用者狀態和卡片盒字典
user_states = {}
user_decks = {}

# 處理訊息事件的函數
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    # 取得使用者資訊
    user_id = event.source.user_id
    user_input = event.message.text

    # 判斷使用者輸入
    if '建立卡片' in user_input:
        text_message = TextSendMessage(text="""歡迎使用建立卡片功能
請選擇您要建立單字卡還是閃卡
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

    elif user_input == '我要建立閃卡':
        # 使用者進入輸入卡片盒名稱的狀態
        user_states[user_id] = 'waiting_for_deck_name'
        reply_text = '請輸入卡片盒名稱'
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)

    elif user_id in user_states and user_states[user_id] == 'waiting_for_deck_name':
        deck_name = user_input
        service_file_path = './client_secret.json'
        spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing'
        all_worksheets = get_all_worksheets(spreadsheet_url, service_file_path, user_id, deck_name)

        if all_worksheets is not None:
            # 提示使用者確認是否使用現有卡片盒
            reply_text = f'已有此卡片盒「{deck_name}」'
            user_states[user_id] = 'waiting_for_confirm_existing_deck'
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
                                "text": f"已有此卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要使用卡片盒「{deck_name}」？",
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

        else:
            # 提示使用者確認是否建立新卡片盒
            reply_text = f'找不到卡片盒「{deck_name}」'
            user_states[user_id] = 'waiting_for_confirm_new_deck'
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
                                "text": f"找不到卡片盒名稱",
                                "weight": "bold",
                                "size": "md"
                            },
                            {
                                "type": "text",
                                "text": f"請問是否要建立卡片盒「{deck_name}」？",
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

    # 使用者確認是否使用現有卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_existing_deck':
        if user_input == '是':
            # 使用者確認使用現有卡片盒
            reply_text = f'確定使用卡片盒「{user_decks[user_id]}」\n請輸入卡片內容'
            # 在這裡處理確定使用現有卡片盒的邏輯，例如將卡片加入現有卡片盒的操作
            user_states[user_id] = 'waiting_for_user_input_content'
        else:
            # 使用者取消使用現有卡片盒
            reply_text = f'取消使用卡片盒「{user_decks[user_id]}」'
        # 回覆使用者
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)
        # 清除用戶狀態
        user_states.pop(user_id, None)

    # 使用者確認是否建立新卡片盒
    elif user_id in user_states and user_states[user_id] == 'waiting_for_confirm_new_deck':
        if user_input == 'Yes':
            # 使用者確認建立新卡片盒
            service_file_path = './client_secret.json'
            spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1diPdtyoqyYGDY7n9pITjU3bMv-i3Crc7OKgFSBhTJNc/edit?usp=sharing'
            # 在這裡處理確定建立新卡片盒的邏輯，例如建立新的工作表
            new_worksheet = create_new_worksheet(user_id, user_decks[user_id], service_file_path, spreadsheet_url)

            # 新工作表建立成功的提示
            reply_text = f'確定建立卡片盒「{user_decks[user_id]}」'
            user_states[user_id] = 'waiting_for_user_input_content'
        else:
            # 使用者取消建立新卡片盒
            reply_text = f'取消建立卡片盒「{user_decks[user_id]}」'
        # 回覆使用者
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)
        # 清除用戶狀態
        user_states.pop(user_id, None)

    else:
        # 其他操作失敗的情況
        reply_text = '操作失敗'
        # 回覆使用者
        message = TextSendMessage(text=reply_text)
        line_bot_api.reply_message(event.reply_token, message)

#主程式
import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)