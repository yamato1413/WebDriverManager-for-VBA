Attribute VB_Name = "WebDriverManager4SeleniumBasic"
Option Explicit

Enum BrowserName
    Chrome
    Edge
End Enum


'// ファイルダウンロード用のWin32API
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If


Private Property Get fso() 'As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property




'// ダウンロードしたWebDriverのzipのデフォルトパス
Public Property Get ZipPath(browser As BrowserName) As String
    Dim path_download As String
    path_download = CreateObject("Shell.Application").Namespace("shell:Downloads").self.path
    Select Case browser
    Case BrowserName.Chrome
        ZipPath = path_download & "\chromedriver_win32.zip"
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: ZipPath = path_download & "\edgedriver_win64.zip"
            Case Else: ZipPath = path_download & "\edgedriver_win32.zip"
        End Select
    End Select
End Property


'// WebDriverの実行ファイルの保存場所

Public Property Get WebDriverPath(browser As BrowserName) As String
    Dim path_AppDataLocal As String
    path_AppDataLocal = CreateObject("Shell.Application").Namespace("shell:Local AppData").self.path
    Select Case browser
        Case BrowserName.Chrome: WebDriverPath = path_AppDataLocal & "\SeleniumBasic\chromedriver.exe"
        Case BrowserName.Edge:   WebDriverPath = path_AppDataLocal & "\SeleniumBasic\edgedriver.exe"
    End Select
End Property



'// ブラウザのバージョンをレジストリから読み取る
'// 出力例　"94.0.992.31"
Public Property Get BrowserVersion(browser As BrowserName)
    Dim reg_version As String
    Select Case browser
        Case BrowserName.Chrome: reg_version = "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\BLBeacon\version"
        Case BrowserName.Edge:   reg_version = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\BLBeacon\version"
    End Select
    
    On Error GoTo Catch
    BrowserVersion = CreateObject("WScript.Shell").RegRead(reg_version)
    Exit Property
    
Catch:
    Err.Raise 4000, , "バージョン情報が取得できませんでした。ブラウザがインストールされていません。"
End Property
'// 出力例　"94"
Public Property Get BrowserVersionToMajor(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToMajor = vers(0)
End Property
'// 出力例　"94.0"
Public Property Get BrowserVersionToMinor(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToMinor = Join(Array(vers(0), vers(1)), ".")
End Property
'// 出力例　"94.0.992"
Public Property Get BrowserVersionToBuild(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToBuild = Join(Array(vers(0), vers(1), vers(2)), ".")
End Property


'// OSが64Bitかどうかを判定する
Public Property Get Is64BitOS() As Boolean
    Dim arch As String
    '戻り値 "AMD64","IA64","x86"のいずれか
    arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE")
    '64bitOSで32bitOfficeを実行している場合、PROCESSOR_ARCHITEW6432に本来のOSのbit数が退避されているので確認
    If InStr(arch, "64") = 0 Then arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITEW6432")
    Is64BitOS = InStr(arch, "64")
End Property




'// 第3引数を省略すれば、ダウンロードフォルダにダウンロードされる
'//     DownloadWebDriver Edge, "94.0.992.31"
'//
'// 第2引数にBrowserVersionプロパティを使えば、現在のブラウザに適合したWebDriverをダウンロードできる
'//     DownloadWebDriver Edge, BrowserVersion(Edge)
'//
'// 第3引数にてパスを指定すれば任意の場所に任意の名前で保存できる。
'//     DownloadWebDriver Edge, "94.0.992.31", "C:\Users\yamato\Desktop\edgedriver_94.zip"
Public Function DownloadWebDriver(browser As BrowserName, ver_webdriver As String, Optional path_save_to As String) As String
    Dim url As String
    Select Case browser
    Case BrowserName.Chrome
        url = Replace("https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip", "{version}", ver_webdriver)
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip", "{version}", ver_webdriver)
            Case Else: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win32.zip", "{version}", ver_webdriver)
        End Select
    End Select
    
    If path_save_to = "" Then path_save_to = ZipPath(browser)   'デフォは"C:Users\USERNAME\Downloads\~~~.zip"
    
    DeleteUrlCacheEntry url
    Dim ret As Long
    ret = URLDownloadToFile(0, url, path_save_to, 0, 0)
    If ret <> 0 Then Err.Raise 4001, , "ダウンロード失敗 : " & url
    
    DownloadWebDriver = path_save_to
End Function



'// zipから中身を取り出して指定の場所に実行ファイルを展開する
'// chromedriver.exe(デフォルトの名前)があるところにchromedriver_94.exeとかで展開できるよう、
'// 元の実行ファイルを上書きしないように一度tempフォルダを作ってから実行ファイルを目的のパスへ移す
'// 普通zipを展開するときは展開先のフォルダを指定するが、この関数はWebDriverの実行ファイルのパスで指定するので注意！(展開するのもexeだけ)
'// 使用例
'//     Extract "C:\Users\yamato\Downloads\chromedriver_win32.zip", "C:\Users\yamato\Downloads\chromedriver_94.exe"
Sub Extract(path_zip As String, path_save_to As String)
    Debug.Print "zipを展開します"
    
    Dim folder_temp
    folder_temp = fso.BuildPath(fso.GetParentFolderName(path_save_to), fso.GetTempName)
    fso.CreateFolder folder_temp
    Debug.Print "    一時フォルダ : " & folder_temp
    
    'PowerShellを使って展開するとマルウェア判定されたので，
    'MS非推奨だがShell.Applicationを使ってzipを解凍する
    
    On Error GoTo Catch
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    'zipファイルに入っているファイルを指定したフォルダーにコピーする
    '文字列を一度()で評価してからNamespaceに渡さないとエラーが出る
    sh.Namespace((folder_temp)).CopyHere sh.Namespace((path_zip)).Items
    
    Dim path_exe As String
    path_exe = fso.BuildPath(folder_temp, Dir(folder_temp & "\*.exe"))
    
    If fso.FileExists(path_save_to) Then fso.DeleteFile path_save_to
    fso.CopyFile path_exe, path_save_to, True
    
    fso.DeleteFolder folder_temp
    Debug.Print "    展開 : " & path_save_to
    Debug.Print "WebDriverを配置しました"
    Exit Sub
    
Catch:
    fso.DeleteFolder folder_temp
    Err.Raise 4002, , "Zipの展開に失敗しました。原因：" & Err.Description
    Exit Sub
End Sub


'// 基本的にはブラウザのバージョンと全く同じバージョンのWebDriverをダウンロードすればいいのだが、
'// ChromeDriverはビルド番号までのバージョンを投げるとおすすめバージョンを教えてくれるらしい？
'// よくわかんないけど、サイトにそう書いてあった。→　https://chromedriver.chromium.org/downloads/version-selection
'// バグフィックスをリリースするから必ずしも一致するとは限らないとか。
Public Function RequestWebDriverVersion(ver_chrome)
    Dim http 'As XMLHTTP60
    Dim url As String
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ver_chrome
    http.Open "GET", url, False
    http.send
    
    If http.statusText <> "OK" Then
        Err.Raise 4003, "サーバーへの接続に失敗しました"
        Exit Function
    End If

    RequestWebDriverVersion = http.responseText
End Function


'// 自動でブラウザのバージョンに一致するWebDriverをダウンロードし、zipを展開、WebDriverのexeを特定のフォルダに配置する
'// デフォルトではC:\Users\USERNAME\Downloadsにダウンロードし、
'// C:\Users\USERNAME\AppData\SeleniumBasic\chromedriver.exe[edgedriver.exe]に配置する
'// 第2引数を指定すれば任意のフォルダ・ファイル名にしてインストールできる
'// 指定したパスの途中のフォルダが存在しなくても、自動で作成する
'// 使用例
'//     InstallWebDriver Chrome, "C:\Users\USERNAME\Desktop\a\b\c\chromedriver_94.exe"
'//     ↑デスクトップに\a\b\c\フォルダが作成されてその中にドライバが配置される
Public Sub InstallWebDriver(browser As BrowserName, Optional path_driver As String)
    Debug.Print "WebDriverをインストールします......"
    
    Dim ver_browser   As String
    Dim ver_webdriver As String
    ver_browser = BrowserVersion(browser)
    Select Case browser
        Case BrowserName.Chrome: ver_webdriver = RequestWebDriverVersion(BrowserVersionToBuild(browser))
        Case BrowserName.Edge:   ver_webdriver = ver_browser
    End Select
    
    Debug.Print "   ブラウザ          : Ver. " & ver_browser
    Debug.Print "   適合するWebDriver : Ver. " & ver_webdriver
    
    Dim path_zip As String
    path_zip = DownloadWebDriver(browser, ver_webdriver)
    
    Do Until fso.FileExists(ZipPath(browser))
        DoEvents
    Loop
    Debug.Print "   ダウンロード完了:" & path_zip
    
    If path_driver = "" Then path_driver = WebDriverPath(browser)
    
    If Not fso.FolderExists(fso.GetParentFolderName(path_driver)) Then
        Debug.Print "   WebDriverの保存先フォルダを作成します"
        CreateFolderEx fso.GetParentFolderName(path_driver)
    End If
    
    Extract path_zip, path_driver
    
    Debug.Print "インストール完了"
End Sub



'// パスに含まれる全てのフォルダの存在確認をしてフォルダを作る関数
'// 使用例
'// CreateFolderEx "C:\a\b\c\d\e\"
Public Sub CreateFolderEx(path_folder As String)
    '// 親フォルダが遡れなくなるところまで再帰で辿る
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    '// 途中の存在しないフォルダを作成しながら降りてくる
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub



'// SeleniumBasicの Driver.Startをこれに置き換えれば、バージョンアップや新規PCへの配布時に余計な操作がいらない
Public Sub SafeOpen(Driver As Selenium.WebDriver, browser As BrowserName)
    
    If Not IsOnline Then Err.Raise 4005, , "オフラインです。インターネットに接続してください。": Exit Sub
    
    '// アップデート処理
    If Not IsLatestDriver(browser) Then
        Dim driver_temp As String
        If fso.FileExists(WebDriverPath(browser)) Then driver_temp = BuckupTempDriver(browser)
        
        Call InstallWebDriver(browser)
    End If
    
    On Error Resume Next
    Select Case browser
        Case BrowserName.Chrome: Driver.Start "chrome"
        Case BrowserName.Edge: Driver.Start "edge"
    End Select
    
    Dim OK As Boolean: OK = Err.Number = 0
    Dim err_number As Long: err_number = Err.Number
    Dim err_desc As String: err_desc = Err.Description
    On Error GoTo 0
    
    If OK Then
        If driver_temp <> "" Then Call DeleteTempDriver(driver_temp)
    Else
        If driver_temp <> "" Then Call RestoreTempDriver(driver_temp, browser)
        Err.Raise err_number, , err_desc
    End If
    
End Sub



'// PCがオンラインかどうかを判定する
'// リクエスト先がgooglなのは障害でページが開けないということは少なそうなので
Public Function IsOnline() As Boolean
    Dim http
    Dim url As String
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://www.google.co.jp/"
    http.Open "GET", url, False
    http.send
    
    Select Case http.statusText
        Case "OK": IsOnline = True
        Case Else: IsOnline = False
    End Select
End Function


'// ドライバーのバージョンを調べる
Function DriverVersion(browser As BrowserName) As String
    If Not fso.FileExists(WebDriverPath(browser)) Then DriverVersion = "": Exit Function
    
    Dim ret As String
    ret = CreateObject("WScript.Shell").Exec(WebDriverPath(browser) & " -version").StdOut.ReadLine
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "\d+\.\d+\.\d+\.\d+"
    DriverVersion = reg.Execute(ret)(0).value
End Function

'// 最新のドライバーがインストールされているか調べる
Function IsLatestDriver(browser As BrowserName) As Boolean
    Select Case browser
    Case BrowserName.Edge
        IsLatestDriver = BrowserVersion(Edge) = DriverVersion(Edge)
    
    '// Chromeは末尾のバージョンがブラウザとドライバーで異なることがある
    Case BrowserName.Chrome
        IsLatestDriver = RequestWebDriverVersion(BrowserVersionToBuild(Chrome)) = DriverVersion(Chrome)
        
    End Select
End Function

'// WebDriverを一時フォルダに退避させる
Function BuckupTempDriver(browser As BrowserName) As String
    Dim folder_temp As String
    folder_temp = fso.BuildPath(fso.GetParentFolderName(WebDriverPath(browser)), fso.GetTempName)
    fso.CreateFolder folder_temp
    
    Dim path_driver As String
    path_driver = fso.BuildPath(folder_temp, "\webdriver.exe")
    fso.MoveFile WebDriverPath(browser), path_driver
    
    BuckupTempDriver = path_driver
End Function

'// WebDriverを一時フォルダからWebDriver置き場にコピーする
Sub RestoreTempDriver(path As String, browser As BrowserName)
    fso.CopyFile path, WebDriverPath(browser), True
    fso.DeleteFolder fso.GetParentFolderName(path)
End Sub

Sub DeleteTempDriver(path As String)
    fso.DeleteFolder fso.GetParentFolderName(path)
End Sub

