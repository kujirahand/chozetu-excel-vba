import urllib.request

# Web APIからXML形式で東京の天気を得るURL --- (*1)
url = 'https://api.aoikujira.com/tenki/week.php?fmt=xml&city=319'

# URLの内容を取得する --- (*2)
req = urllib.request.Request(url)
with urllib.request.urlopen(req) as res:
    # データを読み込む --- (*3)
    bin = res.read()
    # UTF-8で得る --- (*4)
    txt = bin.decode("utf-8")
    # 内容を表示 --- (*5)
    print(txt)

