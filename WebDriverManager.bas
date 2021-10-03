Attribute VB_Name = "WebDriverManager"
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

'// ダウンロードしたWebDriverのzipの保存場所
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


'// WebDriverの実行ファイルの保存場所をレジストリに記録している。
'// デフォルトはドキュメントフォルダ
'// このパスを書き換えるプロシージャは以下の通り
'// ・Property Let WebDriverPath
'// ・InstallWebDriver
Public Property Let WebDriverPath(browser As BrowserName, path_driver As String)
    Select Case browser
        Case BrowserName.Chrome: SaveSetting "WebDriverManager", "WebDriverPath", "Chrome", path_driver
        Case BrowserName.Edge:   SaveSetting "WebDriverManager", "WebDriverPath", "Edge", path_driver
    End Select
End Property
Public Property Get WebDriverPath(browser As BrowserName) As String
    Select Case browser
        Case BrowserName.Chrome: WebDriverPath = GetSetting("WebDriverManager", "WebDriverPath", "Chrome", "C:" & Environ("HOMEPATH") & "\Documents\WebDriver\Chrome\chromedriver.exe")
        Case BrowserName.Edge:   WebDriverPath = GetSetting("WebDriverManager", "WebDriverPath", "Edge", "C:" & Environ("HOMEPATH") & "\Documents\WebDriver\Edge\msedgedriver.exe")
    End Select
End Property




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
Public Property Get BrowserVersionToMajor(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToMajor = vers(0)
End Property
Public Property Get BrowserVersionToMinor(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToMinor = Join(Array(vers(0), vers(1)), ".")
End Property
Public Property Get BrowserVersionToBuild(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToBuild = Join(Array(vers(0), vers(1), vers(2)), ".")
End Property



Public Property Get Is64BitOS() As Boolean
    Dim arch As String
    arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE")
    Is64BitOS = CBool(InStr(arch, "64"))
End Property



'// 第3引数にてパスを指定すれば任意の場所に任意の名前で保存できる。
'// 使用例 DownloadWebDriver Edge, "94.0.992.31", "C:\Users\yamato\Desktop\edge.zip"
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
'// chromedriver.exe(デフォルトの名前)があるところにchromedriver_94.exeとかで展開できるように、
'// 元の実行ファイルを上書きしないように一度tempフォルダを作ってから実行ファイルを目的のパスへ移す
'// 使用例 Extract "C:\Users\yamato\Downloads\chromedriver_win32.zip","C:\Users\yamato\Downloads\chromedriver_94.exe"
Sub Extract(path_zip As String, path_save_to As String)
    Dim folder_temp
    folder_temp = fso.BuildPath(fso.GetParentFolderName(path_save_to), fso.GetTempName)
    fso.CreateFolder folder_temp
    
    'Shell.Applicationを使う方法はMS非推奨らしいのでPowerShellで展開する
    Dim command As String, ex As Object 'WshExec
    command = "Expand-Archive -Path " & path_zip & " -DestinationPath " & folder_temp & " -Force"
    Set ex = wsh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & command)
    
    '// コマンド失敗時
    If ex.Status = WshFailed Then: Err.Raise 4002, , "Zipの展開に失敗しました": Exit Sub
    
    Do While ex.Status = 0 'WshRunning
        DoEvents
    Loop
    
    Dim path_exe_from As String, path_exe_to As String
    path_exe_from = fso.BuildPath(folder_temp, Dir(folder_temp & "\*.exe"))
    
    fso.MoveFile path_exe_from, path_save_to
    fso.DeleteFolder folder_temp
End Sub



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



'パスに含まれる全てのフォルダの存在確認をしてフォルダを作る関数
Sub CreateFolderEx(path_folder As String)
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub



'// TinySeleniumVBA専用の関数です。SeleniumBasicでは動きません。
'//
Sub SafeOpen(Driver As WebDriver, browser As BrowserName, Optional path_driver As String)
    If path_driver = "" Then path_driver = WebDriverPath(browser)
    
    If Not fso.FileExists(WebDriverPath(browser)) Then
        Debug.Print "WebDriverが見つかりません"
        InstallWebDriver browser
    End If
    
    Select Case browser
        Case BrowserName.Chrome: Driver.Chrome path_driver
        Case BrowserName.Edge:   Driver.Edge path_driver
    End Select
    
    On Error GoTo ErrHandler
    Driver.OpenBrowser
    Exit Sub
    
ErrHandler:
    Driver.Shutdown
    InstallWebDriver browser
    Resume
End Sub




