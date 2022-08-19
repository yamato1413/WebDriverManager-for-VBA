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
Dim Driver As WebDriver
Driver.Edge "WebDriverへのパス"
Driver.OpenBrowser
'   ↓
SafeOpen Driver, Edge [,"WebDriverへのパス"] '// 第3引数は省略可
```
```VB
'// SeleniumBasic
Dim Driver As Selenium.ChromeDriver
Driver.Start 
'   ↓
SafeOpen Driver, Chrome [,"WebDriverへのパス"] '// 第3引数は省略可
```

この```SafeOpen```は、ブラウザを開く前にWebDriverの存在を確認し、なければWebDriverのダウンロード・展開を開始します。
また、```Driver.OpenBrowser[Start]```がコケた場合(WebDriverとブラウザのバージョンが違う時)、適合するWebDriverをダウンロード・展開し```Driver.OpenBrowser[Start]```をリトライします。

つまり、```SafeOpen```でマクロを書いておけば、バージョンアップ時どころか、マクロ配布時にWebDriverを同梱したりWebDriverの入れ方マニュアルを作らなくてよくなります。
これはマクロ開発者にとって非常にうれしいことだと思います。

以下はSampleコードです。

```VB
'// TinySeleniumVBA
Option Explicit

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
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub Sample()
    Dim Driver As New Selenium.ChromeDriver
    SafeOpen Driver, Chrome
    Driver.Get "https://www.google.co.jp/?q=selenium"
    Sleep 3000
    Driver.Close
End Sub
```

#### 補足
WebDriverの存在を確認すると書きましたが、実際にはどこを確認しているのか。
デフォルトでは以下の場所を確認しています。

```
TinySeleniumVBA版
    C:\Users\USERNAME\Documents\WebDriver\edgedriver.exe[chromedriver.exe]
SeleniumBasic版
    C:\Users\USERNAME\AppData\Local\SeleniumBasic\edgedriver.exe[chromedriver.exe]
    C:\Program Files\SeleniumBasic\edgedriver.exe[chromedriver.exe]
    C:\Program Files (x86)\SeleniumBasic\edgedriver.exe[chromedriver.exe]
    のいずれか
```

最初の例で以下のように書きました

```VB
SafeOpen Driver, Edge [,"WebDriverへのパス"] '// 第3引数は省略可
```

WebDriverを保存する場所にこだわりがあるなら引数で指定してもいいですが，
パスを省略した場合は上記のデフォルトパスを確認してWebDriverが存在しなければ自動でインストールを始めるので、
特にデフォルトのパスに異論がなければ

```VB
SafeOpen Driver, Edge
```

で十分です。

よいスクレイピングライフを！
