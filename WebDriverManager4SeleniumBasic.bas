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
Private Declare  Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare  Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If


Private Property Get fso() As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property
Private Property Get wsh() 'As WshShell
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("WScript.Shell")
    Set wsh = obj
End Property




'// ダウンロードしたWebDriverのzipのデフォルトパス
Private Property Get ZipPath(browser As BrowserName) As String
    Select Case browser
        Case BrowserName.Chrome
            ZipPath = "C:" & Environ("HOMEPATH") & "\Downloads\chromedriver_win32.zip"
        Case BrowserName.Edge
            Select Case Is64BitOS
                Case True: ZipPath = "C:" & Environ("HOMEPATH") & "\Downloads\edgedriver_win64.zip"
                Case Else: ZipPath = "C:" & Environ("HOMEPATH") & "\Downloads\edgedriver_win32.zip"
            End Select
    End Select
End Property


'// WebDriverの実行ファイルの保存場所をレジストリに記録している
'// デフォルトはドキュメントフォルダ（SeleniumBasicがインストールされている場合は\AppData\Local\SeleniumBasic\）
'// このパスを書き換えるプロシージャは以下の通り
'//     chromedriver.exe
'//     Property Let WebDriverPath
'//     InstallWebDriver
Public Property Let WebDriverPath(browser As BrowserName, path_driver As String)
    Select Case browser
        Case BrowserName.Chrome: SaveSetting "WebDriverManager", "WebDriverPath", "Chrome", path_driver
        Case BrowserName.Edge:   SaveSetting "WebDriverManager", "WebDriverPath", "Edge", path_driver
    End Select
End Property
Public Property Get WebDriverPath(browser As BrowserName) As String
    Select Case browser
        Case BrowserName.Chrome: WebDriverPath = GetSetting("WebDriverManager", "WebDriverPath", "Chrome", "C:" & Environ("HOMEPATH") & "\AppData\Local\SeleniumBasic\chromedriver.exe")
        Case BrowserName.Edge:   WebDriverPath = GetSetting("WebDriverManager", "WebDriverPath", "Edge", "C:" & Environ("HOMEPATH") & "\AppData\Local\SeleniumBasic\edgedriver.exe")
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
    
    Dim version As String
    version = CreateObject("WScript.Shell").RegRead(reg_version)
    
    If version = "" Then Err.Raise 4000, , "バージョン情報が取得できませんでした"
    BrowserVersion = version
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
    arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE") '戻り値 "AMD64","IA64","x86"のいずれか
    Is64BitOS = CBool(InStr(arch, "64"))
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
    'Shell.Applicationを使う方法はMS非推奨らしいのでPowerShellで展開する
    Dim command As String, ex As Object 'WshExec
    command = "Expand-Archive -Path " & path_zip & " -DestinationPath " & folder_temp & " -Force"
    Set ex = wsh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & command)
    
    '// コマンド失敗時
    If ex.Status = WshFailed Then GoTo Catch
    
    Do While ex.Status = 0 'WshRunning
        DoEvents
    Loop
    
    On Error GoTo Catch
    Dim path_exe_from As String, path_exe_to As String
    path_exe_from = fso.BuildPath(folder_temp, Dir(folder_temp & "\*.exe"))
    
    fso.MoveFile path_exe_from, path_save_to
    fso.DeleteFolder folder_temp
    Debug.Print "    展開 : " & path_save_to
    Debug.Print "WebDriverを配置しました"
    Exit Sub
Catch:
    fso.DeleteFolder folder_temp
    Err.Raise 4002, , "    Zipの展開に失敗しました"
End Sub


'// 基本的にはブラウザのバージョンと全く同じバージョンのWebDriverをダウンロードすればいいのだが、
'// ChromeDriverはビルド番号までのバージョンを投げるとおすすめバージョンを教えてくれるらしい？
'// よくわかんないけど、サイトにそう書いてあった。→　https://chromedriver.chromium.org/downloads/version-selection
'// バグフィックスをリリースするから必ずしも一致するとは限らないとか。
Function RequestWebDriverVersion(ver_chrome)
    Dim http 'As XMLHTTP60
    Dim url As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "http://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ver_chrome
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
'// C:\Users\USERNAME\Documents\WebDriver\Chrome[Edge]\chromedriver.exe[msedgedriver.exe]に配置する
'// 第2引数を指定すれば任意のフォルダ・ファイル名にしてインストールできる
'// 指定したパスの途中のフォルダが存在しなくても、自動で作成する
'// 使用例
'//     InstallWebDriver Chrome, "C:\Users\USERNAME\Desktop\a\b\c\chromedriver_94.exe"
'//     ↑デスクトップに\a\b\c\フォルダが作成されてその中にドライバが配置される
Sub InstallWebDriver(browser As BrowserName, Optional path_driver As String)
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
    
    If path_driver <> "" Then WebDriverPath(browser) = path_driver
    
    If Not fso.FolderExists(fso.GetParentFolderName(WebDriverPath(browser))) Then
        Debug.Print "   WebDriverの保存先フォルダを作成します"
        CreateFolderEx fso.GetParentFolderName(WebDriverPath(browser))
    End If
    
    Extract path_zip, WebDriverPath(browser)
    
    Debug.Print "インストール完了"
End Sub



'// パスに含まれる全てのフォルダの存在確認をしてフォルダを作る関数
'// 使用例
'// CreateFolderEx "C:\a\b\c\d\e\"
Sub CreateFolderEx(path_folder As String)
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub


'// WebDriverの存在チェックをして無ければインストールする
'// また、WebDriverが存在してもバージョン不一致でブラウザが開けなかった場合もWebDriverを再インストールする
'// TinySeleniumVBAの "Driver.Chrome[Edge] path" と "Driver.OpenBrowser"をこれに置き換えれば、
'// バージョンアップや新規PCへの配布時に余計な操作がいらない
Sub SafeOpen(Driver As Selenium.WebDriver, browser As BrowserName, Optional path_driver As String)
    If path_driver = "" Then path_driver = WebDriverPath(browser)
    
    If Not fso.FileExists(path_driver) Then
        Debug.Print "WebDriverが見つかりません"
        InstallWebDriver browser
    End If
    
    On Error GoTo Catch
    Dim counter_try As Long
    Select Case browser
        Case BrowserName.Chrome: Driver.Start
        Case BrowserName.Edge:   Driver.Start "edge"
    End Select
    Exit Sub
    
Catch:
    counter_try = counter_try + 1
    Driver.Close
    If counter_try > 1 Then Err.Raise 4004, , "ブラウザのオープンに失敗しました"
    InstallWebDriver browser
    Resume
End Sub






