#from _future_ import print_function
#from apiclient.discovery import build
#from httplib2 import Http
#from oauth2client import file, client, tools

#import time
#import re
#import datetime
#import random
#import codecs
#import sys
#import json

import numpy as np
import pandas as pd
from ExchangeCrawler import TaiwanExchangeCrawler
import utils

from flask import Flask, request, abort
from urllib.request import urlopen
#from oauth2client.service_account import ServiceAccountCredentials

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError,LineBotApiError
)

################################

from linebot.models import *

app = Flask(__name__)

# Channel Access Token
line_bot_api = LineBotApi('/kNecMdZomV7O0z9Ie8o8y5BFzmpbEqGLfWNrh74iA7qmRdPUdwM49TbvfL8/Tfr8QiFyfpNWA+rhsYjvknQHkx6Btj+wc2nn0Zx7bqNwQt9760NOK2WxrsQzkBLCHmf2SL1jTX3iP+SLbZGSFCmEAdB04t89/1O/w1cDnyilFU=')
# Channel Secret
handler = WebhookHandler('e571aa49803b3c0444552238b7ae1004')

crawler = TaiwanExchangeCrawler()

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

# 處理訊息
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    print(event)
    text=event.message.text
    if text != 'all' :
        names = text.split(',')
    else :
        names = text
    #init crawler
    df_CashBuy, df_CashSell, df_SpotBuy, df_SpotSell, errorlog = crawler.GetBest(banklist=names)
    utils.render_mpl_table(df_SpotBuy, header_columns=0)
    imglink_CashBuy  = utils.GetImageLink(df_CashBuy)
    imglink_CashSell = utils.GetImageLink(df_CashSell)
    imglink_SpotBuy  = utils.GetImageLink(df_SpotBuy)
    imglink_SpotSell = utils.GetImageLink(df_SpotSell)

    Image_Carousel = TemplateSendMessage(alt_text='Exchange Rate',
                        template=ImageCarouselTemplate(
                            columns=[
                                ImageCarouselColumn(image_url=imglink_CashBuy, action=URITemplateAction(
                                                    label='CashBuy',  uri=imglink_CashBuy)),
                                ImageCarouselColumn(image_url=imglink_CashSell, action=URITemplateAction(
                                                    label='CashSell', uri=imglink_CashSell)),
                                ImageCarouselColumn(image_url=imglink_SpotBuy, action=URITemplateAction(
                                                    label='SpotBuy',  uri=imglink_SpotBuy)),
                                ImageCarouselColumn(image_url=imglink_SpotSell, action=URITemplateAction(
                                                    label='SpotSell', uri=imglink_SpotSell))
                            ]
                        )
                    )
    line_bot_api.reply_message(event.reply_token, Image_Carousel)
    # message1 = ImageSendMessage(original_content_url=imglink_SpotBuy, preview_image_url=imglink_SpotBuy)
    # message2 = ImageSendMessage(original_content_url=imglink_SpotSell, preview_image_url=imglink_SpotSell)
    # line_bot_api.reply_message(event.reply_token, [message1, message2])


import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)