# Python標準のダイアログ命令の利用を宣言
import tkinter as tk
import tkinter.messagebox as mb
tk.Tk().withdraw()

# コマンドライン引数をチェック
import sys
if len(sys.argv) < 2:
    mb.showinfo("error","引数が足りません")
    quit()
# コマンドライン引数の一つ目をメッセージボックスに表示
mb.showinfo("Pythonのダイアログ", sys.argv[1])

