# WebDriverManager-for-VBA

## インストール

TinySeleniumVBAを使ってる人はWebDriverManager4TinySelenium.basを、
SeleniumBasicを使っている人はWebDriverManager4SeleniumBasic.basを
インポートするだけです。


gitが分かる人は
```
git clone github.com/yamato1413/WebDriverManager-for-VBA
```
でダウンロードしてもらえればいいですし，分からない人は右上の「Code」という緑のボタンからZIPでダウンロードしてください。

## SafeOpen関数
ブラウザを開く処理を```SafeOpen```に書き換えることで、WebDriverのバージョンを一切気にする必要がなくなります。

```VB
'//TinySeleniumVBA
Dim Driver As New WebDriver
Driver.Edge "WebDriverへのパス"
Driver.OpenBrowser
'   ↓
SafeOpen Driver, Edge [,"WebDriverへのパス"] '// 第3引数は省略可
```
```VB
'//SeleniumWrapper
Dim Driver As New WebDriver
Driver.Edge "WebDriverへのパス"
Driver.OpenBrowser
'   ↓
SafeOpen Driver, BN_Edge [,"WebDriverへのパス"] '// 第3引数は省略可
```
```VB
'// SeleniumBasic
Dim Driver As New Selenium.ChromeDriver
Driver.Start 
'   ↓
Dim Driver As New Selenium.WebDriver
SafeOpen Driver, Chrome [,"WebDriverへのパス"] '// 第3引数は省略可
```

ブラウザを開く前にWebDriverの存在・バージョンをチェックします。
WebDriverが存在しない、またはバージョンがブラウザと異なる場合にWebDriverのダウンロード・展開を開始します。

存在しない時にもダウンロードを行うので、バージョンアップ時だけでなく、マクロ配布時にWebDriverを同梱したりWebDriverの入れ方マニュアルを作らなくてよくなります。

以下はSampleコードです。

```VB
'// TinySeleniumVBA
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub Sample()
    Dim Driver As New WebDriver
    SafeOpen Driver, Edge
    Driver.Navigate "https://www.google.co.jp/?q=selenium"
    Sleep 3000
    Driver.ShutDown
End Sub
```
```VB
'// SeleniumBasic
Public Sub Sample()
    Dim Driver As New Selenium.ChromeDriver
    SafeOpen Driver, Chrome
    Driver.Get "https://www.google.co.jp/?q=selenium"
    Driver.Wait 3000
    Driver.Quit
End Sub
```

よいスクレイピングライフを！
