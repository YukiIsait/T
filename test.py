# -*- coding: UTF-8 -*-
import requests as req
import json
import sys
import time

path = sys.path[0] + r'/test.txt'
num1 = 0

# Line: 13 插入id
# Line: 15 插入key





# End 插入结束
print(id)
print(secret)
api_list = [
    r'https://graph.microsoft.com/v1.0/me/drive/root',
    r'https://graph.microsoft.com/v1.0/me/drive',
    r'https://graph.microsoft.com/v1.0/drive/root',
    r'https://graph.microsoft.com/v1.0/users',
    r'https://graph.microsoft.com/v1.0/me/messages',
    r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    r'https://graph.microsoft.com/v1.0/me/drive/root/children',
    r'https://graph.microsoft.com/v1.0/me/mailFolders',
    r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories'
]

def gettoken(refresh_token):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'client_id': id,
            'client_secret': secret,
            'redirect_uri': 'http://localhost:53682/'}
    html = req.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token


def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    access_token = gettoken(refresh_token)
    headers = {'Authorization': access_token,
               'Content-Type': 'application/json'}
    try:
        for index in range(len(api_list)):
            print(index, '调用结果:', req.get(api_list[index], headers=headers).status_code, api_list[index])
        print('此次运行结束时间为:', time.asctime(time.localtime(time.time())))
    except Exception as e:
        print(e)


for _ in range(3):
    main()
