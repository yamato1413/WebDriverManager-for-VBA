# WebDriverManager-for-VBA

TinySeleniumVBAを使ってる人はWebDriverManager4TinySelenium.basを、SeleniumBasicを使っている人はWebDriverManager4SeleniumBasic.basをインポートするだけです。

## SafeOpen関数
ブラウザを開く処理を```SafeOpen```に書き換えることで、WebDriverのバージョンを一切気にする必要がなくなります。

```VB:TinySeleniumVBA
Dim Driver As WebDriver
Driver.Edge  "path\to\edgedriver.exe"
Driver.OpenBrowser
'   ↓
SafeOpen Edge, "path\to\edgedriver.exe" '// 第2引数は書かなくてもOK
```
```VB:SeleniumBasic
Dim Driver As Selenium.ChromeDriver
Driver.Start 
'   ↓
SafeOpen Chrome
```

この```SafeOpen```は、ブラウザを開く前にWebDriverの存在を確認し、なければWebDriverのダウンロード・展開を開始します。
また、```Driver.OpenBrowser[Start]```がコケた場合(WebDriverとブラウザのバージョンが違う時)、適合するWebDriverをダウンロード・展開し```Driver.OpenBrowser[Start]```をリトライします。、

つまり、```SafeOpen```でマクロを書いておけば、バージョンアップ時どころか、マクロ配布時にWebDriverを同梱したりWebDriverの入れ方マニュアルを作らなくてよくなります。
これはマクロ開発者にとって非常にうれしいことだと思います。

以下はSampleコードです。

```VB:TinySeleniumVBA
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
```VB:SeleniumBasic
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
```

最初の例で

```VB
SafeOpen Edge, "path\to\edgedriver.exe" '// 第2引数は書かなくてもOK
```

と書きましたが、実際には上記のパスを確認してWebDriverが存在しなければ勝手にインストールが始まるので、特にデフォルトのパスに異論がなければ

```VB
SafeOpen Driver, Edge
```

で十分です。

##その他関数
WebDriverManager-for-VBAには```SafeOpen```以外にも```InstallWebDriver``` ```DownloadWebDriver``` 
 ```BrowserVersion```などの関数が用意されています。というか```SafeOpen```はこれらをラップしただけです。

```VB
'// 自動でブラウザのバージョンに一致するWebDriverをダウンロードし、
'// zipを展開、WebDriverのexeを特定のフォルダに配置する
'// デフォルトではC:\Users\USERNAME\Downloadsにダウンロードし、
'// C:\Users\USERNAME\Documents\WebDriver\chromedriver.exe[edgedriver.exe]に配置する
InstallWebDriver Edge
'// 第2引数を指定すれば任意のフォルダ・ファイル名にしてインストールできる
'// 指定したパスの途中のフォルダが存在しなくても、自動で作成する
InstallWebDriver Chrome, "C:\Users\USERNAME\Desktop\a\b\c\chromedriver_94.exe"
'//     ↑デスクトップに\a\b\c\フォルダが作成されてその中にドライバが配置される

'---------------------------------------------------------------------------------------------

'// 第3引数を省略すれば、ダウンロードフォルダにダウンロードされる
DownloadWebDriver Edge, "94.0.992.31"
'// 第2引数にBrowserVersionプロパティを使えば、現在のブラウザに適合したWebDriverをダウンロードできる
DownloadWebDriver Edge, BrowserVersion(Edge)
'// 第3引数にてパスを指定すれば任意の場所に任意の名前で保存できる。
DownloadWebDriver Edge, "94.0.992.31", "C:\Users\yamato\Desktop\edgedriver_94.zip"
```

あまり汎用的な関数はないのですが、```CreateFolderEx```は使える子だと思うのでぜひパクってください。
FSOのCreateFolderメソッドは親フォルダまでは存在していないとコケますが、この関数はパスをたどって存在しないフォルダを全部作ってくれます。

```VB
'// パスに含まれる全てのフォルダの存在確認をしてフォルダを作る関数
CreateFolderEx "C:\a\b\c\d\e\"

'// 実装
Sub CreateFolderEx(path_folder As String)
    '// 親フォルダが遡れなくなるところまで再帰で辿る
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    '// 途中の存在しないフォルダを作成しながら降りてくる
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub
```

よいスクレイピングライフを！
