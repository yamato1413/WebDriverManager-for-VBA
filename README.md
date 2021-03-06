# WebDriverManager-for-VBA

### バージョン
2021/10/03 Ver.0.1.0 リリース

2021/10/03 Ver.0.2.0 Tiny版とBasic版を分離

2021/10/04 Ver.0.3.0 Basic版にSafeOpen追加

2021/10/05 Ver.0.3.5 重大バグ修正(FailedとSuccessの条件を取り違え)

2021/10/06 Ver.0.4.0 Basic版SafeOpenを大幅修正

2021/10/12 Ver.0.5.0 Tiny版SafeOpenを大幅修正

2021/10/12 Ver.0.5.1 SafeOpenのミス修正

2021/10/12 Ver.0.6.0 マルウェア判定を受けることがあるのでzipの展開方法を変更しました

2022/04/21 Ver.1.0.0 64Bitの判定を正しくできるようにしました。バージョンエラーが起きるときのみ更新するようにしました。

2022/06/15 Ver.1.0.2 フォルダの取得処理をドライブに依存しない形に書き換えました。

2022/06/15 Ver.1.1.0 WebDriverが最新かどうかのチェックロジックを見直しました。

2022/06/15 Ver.1.1.1 SafeOpen関数内のエラー処理を見直しました。

2022/06/15 Ver.1.1.2 かなり古いバージョンのWebDriverのバージョン取得に失敗する問題を修正しました。


---

- 先日アップデートしたときにFailedとSuccessの条件を取り違えてしまったので，成功していてもzipの展開に失敗しましたというエラー吐いて途中で止まってしまいます。
→Ver.0.3.5にて修正

- SeleniumBasic版においてWebDriverの終了ができずWebDriverの書き換えに失敗し更新できない状態になっていました。 根本的な解決ができず，毎回最新ドライバをインストールする形で対応しました。
→Ver.0.4.0にて修正

- TinySelenium版においてWebDriverの書き換えはできるもののリトライのブラウザオープンに失敗する状態になっていました。 根本的な解決ができず，毎回最新ドライバをインストールする形で対応しました。
→Ver.0.5.0にて修正

- Ver.0.4.0および0.5.0において毎回ドライバをダウンロードしていたのを、バージョンエラーが発生したときのみ更新するように修正しました。
また、リファクタリングしたいところは大量にあるものの、とりあえず修正しなければいけない箇所はひと通り直ったので、今回からバージョンがVer.1.0となります。

- 実行時エラーが発生したときに最新でないという判断をしていたのを、ブラウザオープン前にWebDriverのバージョンを拾ってきて判断するように変更しました。これによりバージョンエラー以外の理由でブラウザが開けなかった時のエラー情報の握り潰しを防ぎます。これに合わせてエラー処理も変更しました。

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
Driver.Edge "path\to\edgedriver.exe"
Driver.OpenBrowser
'   ↓
SafeOpen Driver, Edge, "path\to\edgedriver.exe" '// 第3引数は書かなくてもOK
```
```VB
'// SeleniumBasic
Dim Driver As Selenium.ChromeDriver
Driver.Start 
'   ↓
SafeOpen Driver, Chrome
```

この```SafeOpen```は、ブラウザを開く前にWebDriverの存在を確認し、なければWebDriverのダウンロード・展開を開始します。
また、```Driver.OpenBrowser[Start]```がコケた場合(WebDriverとブラウザのバージョンが違う時)、適合するWebDriverをダウンロード・展開し```Driver.OpenBrowser[Start]```をリトライします。、

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
```

最初の例で以下のように書きました

```VB
SafeOpen Driver, Edge, "path\to\edgedriver.exe" '// 第3引数は書かなくてもOK
```

WebDriverを保存する場所にこだわりがあるなら引数で指定してもいいですが，
パスを省略した場合は上記のデフォルトパスを確認してWebDriverが存在しなければ自動でインストールを始めるので、
特にデフォルトのパスに異論がなければ

```VB
SafeOpen Driver, Edge
```

で十分です。

## その他関数
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
