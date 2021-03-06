# 「ベイズ」でスパムメールを分類しよう！

以下の解説は、p415 5章「文章のスパム判定」に対する捕捉です。
ベイズを利用して、文章の自動分類に挑戦する方法を紹介します。

## 「ベイズ」「ベイジアンフィルタ」とは？

「ベイズの定理」とは、トーマス・ベイズにちなんで名付けられたもので、ある事象に関連する可能性のある条件についての事前の知識に基づいて、その事象の確率を記述するものです。そして「ベイズの定理」を応用したものに、推計統計学の手法の一つである「ベイズ推定」があります。

「ベイズ推定」を利用して対象となるデータを解析・学習し分類する為のフィルタを「ベイジアンフィルタ(Bayesian Filter)」と呼びます。スパム判定（迷惑メール判定）を行うソフトウェアで使われています。

このベイズ推定を応用したスパム判定のプログラムを作ることができます。

## ベイズによるスパム判定を行う方法

なお、書籍の中では「ベイズ」ではなく「SVM」というアルゴリズムを利用して文章の「スパム判定」を行っています。

p425 の学習モデル作成のプログラムを以下のように書き換えます。
すると、「ベイズ」を使うように変更できます。(プログラムを2行変えるだけです。）

```
Sub TrainBayes()
    Dim ModelFile
    ModelFile = ThisWorkbook.Path & "\spam-nb-model.pkl"
    MsgBox SKLTrain(Sheets("学習用データ"), "native_bayes", 1, ModelFile)
End Sub
```

さらに、p426の上部にあるプロシージャ「スパム判定実行」を以下のように書き換えます。

```
Sub スパム判定実行()
    Text = Sheets("スパム判定").Range("A2").Value
    MakeBoWByText Text
    ' ModelFile = ThisWorkbook.Path & "\spam-svm-model.pkl"
    ModelFile = ThisWorkbook.Path & "\spam-nb-model.pkl"
    MsgBox SKLPredict(Sheets("作業用"), ModelFile)
End Sub
```

上記の修正を追加したプログラムを、以下のパスに配置しています。

 - src/ch5/spam-check-bayes.xlsm

参考にしてください。



