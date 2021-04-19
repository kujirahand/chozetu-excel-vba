import urllib.request, json
import pprint

# Web APIからJSONデータを取得 --- (*1)
api = 'https://api.aoikujira.com/tenki/week.php?fmt=json&city=319'
req = urllib.request.Request(api)
with urllib.request.urlopen(req) as res:
    body = json.load(res)
# JSONデータを読みやすく表示 --- (*2)
pprint.pprint(body)
