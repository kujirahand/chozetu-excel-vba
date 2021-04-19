import urllib.request
import json

# APIのURLを指定 --- (*1)
api = 'https://api.aoikujira.com/tenki/week.php?fmt=json&city=319'

# APIからJSONデータを取得 --- (*2)
req = urllib.request.Request(api)
with urllib.request.urlopen(req) as res:
    # JSONデータをPythonのデータ型に読込 --- (*3)
    body = json.load(res)
    # データを表示
    print(body)

