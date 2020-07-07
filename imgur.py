import pyimgur
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from op_advance import Op

# CLIENT_ID = "97e99a6b3293f49"
# client_secret = '0a5d3149f04b54bac5bb86e22a8e0a8667316501'
# access_token = '3c62cc11a75c502d7c6ad8d93fcbaff84c3638d1'
# refresh_token = 'b4a8c62ba98aee2bd2557fa7e654fb7e8d3f459f'

# PATH = "C:\\Users\\RF\\Desktop\\coding\\haha\\TaiwanBank\\Bear.jpg"

# im = pyimgur.Imgur(CLIENT_ID, client_secret, access_token, refresh_token)
# uploaded_image = im.upload_image(PATH, title="Uploaded with PyImgur")
# print(uploaded_image.title)
# print(uploaded_image.link)
# print(uploaded_image.size)
# print(uploaded_image.type)

# help(im.upload_image)
# help(pyimgur.Imgur)

def UploadToImgur(file, title) :
    client_id = '97e99a6b3293f49'
    client_secret = '0a5d3149f04b54bac5bb86e22a8e0a8667316501'
    access_token = '3c62cc11a75c502d7c6ad8d93fcbaff84c3638d1'
    refresh_token = 'b4a8c62ba98aee2bd2557fa7e654fb7e8d3f459f'
    im = pyimgur.Imgur(client_id)
    uploaded_image = im.upload_image(file, title=title)
    return uploaded_image

def GetImageLink() :
    path = 'image.png'
    header, op_dif = Op().analysis_opdata()
    Op().data_visualization(header[0:2], op_dif[0:2], '自營bcbp')
    plt.savefig(path)
    img = UploadToImgur(path, title='update from python')
    # img = UploadToImgur(path, title='update from python')
    print(img.link)
    return img.link


