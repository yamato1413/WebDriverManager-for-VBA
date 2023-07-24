Attribute VB_Name = "WebDriverManager4SeleniumBasic"
Option Explicit

Enum BrowserName
    Chrome
    Edge
End Enum

#Const DEV = 0

'// ファイルダウンロード用のWin32API
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If




#If DEV Then
    Dim fso As New Scripting.FileSystemObject
    Dim wsh As New WshShell
    Dim shell As New Shell32.shell
#Else
    Private Property Get fso() As Object
        Static Obj As Object
        If Obj Is Nothing Then Set Obj = CreateObject("Scripting.FileSystemObject")
        Set fso = Obj
    End Property
    
    Private Property Get wsh() As Object
        Static Obj As Object
        If Obj Is Nothing Then Set Obj = CreateObject("WScript.Shell")
        Set wsh = Obj
    End Property
    
    Private Property Get shell() As Object
        Static Obj As Object
        If Obj Is Nothing Then Set Obj = CreateObject("Shell.Application")
        Set shell = Obj
    End Property
#End If


Public Property Get ZipPath(Browser As BrowserName) As String
    Dim DownloadFolderPath As String
    DownloadFolderPath = shell.Namespace("shell:Downloads").Self.path
    
    Select Case Browser
    Case BrowserName.Chrome
        Select Case Is64BitOS
            Case True: ZipPath = DownloadFolderPath & "\chromedriver-win64.zip"
            Case Else: ZipPath = DownloadFolderPath & "\chromedriver-win32.zip"
        End Select
        
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: ZipPath = DownloadFolderPath & "\edgedriver_win64.zip"
            Case Else: ZipPath = DownloadFolderPath & "\edgedriver_win32.zip"
        End Select
    End Select
End Property


Public Property Get WebDriverPath(Browser As BrowserName) As String
    Dim SeleniumPath1 As String
    Dim SeleniumPath2 As String
    Dim SeleniumPath3 As String
    Dim SeleniumPath4 As String
    
    SeleniumPath1 = Environ("LocalAppData") & "\SeleniumBasic\"
    SeleniumPath2 = Environ("Programfiles(x86)") & "\SeleniumBasic\"
    SeleniumPath3 = Environ("ProgramW6432") & "\SeleniumBasic\"
    SeleniumPath4 = Environ("Programfiles") & "\SeleniumBasic\"
    
    Dim BasePath As String
    Select Case True
        Case fso.FolderExists(SeleniumPath1): BasePath = SeleniumPath1
        Case fso.FolderExists(SeleniumPath2): BasePath = SeleniumPath2
        Case fso.FolderExists(SeleniumPath3): BasePath = SeleniumPath3
        Case fso.FolderExists(SeleniumPath4): BasePath = SeleniumPath4
    End Select
    
    Select Case Browser
        Case BrowserName.Chrome: WebDriverPath = BasePath & "chromedriver.exe"
        Case BrowserName.Edge:   WebDriverPath = BasePath & "edgedriver.exe"
    End Select
End Property

Public Property Get BrowserVersion(Browser As BrowserName)
    Dim EdgePath1 As String
    Dim EdgePath2 As String
    Dim EdgePath3 As String
    Dim ChromePath1 As String
    Dim ChromePath2 As String
    Dim ChromePath3 As String
    EdgePath1 = Environ("Programfiles(x86)") & "\Microsoft\Edge\Application\msedge.exe"
    EdgePath2 = Environ("ProgramW6432") & "\Microsoft\Edge\Application\msedge.exe"
    EdgePath3 = Environ("Programfiles") & "\Microsoft\Edge\Application\msedge.exe"
    ChromePath1 = Environ("Programfiles(x86)") & "\Google\Chrome\Application\chrome.exe"
    ChromePath2 = Environ("ProgramW6432") & "\Google\Chrome\Application\chrome.exe"
    ChromePath3 = Environ("Programfiles") & "\Google\Chrome\Application\chrome.exe"
    
    Dim BrowserFilePath As String
    Dim TargetFile
    Select Case Browser
    Case Edge
        Select Case True
            Case fso.FileExists(EdgePath1): BrowserFilePath = EdgePath1
            Case fso.FileExists(EdgePath2): BrowserFilePath = EdgePath2
            Case fso.FileExists(EdgePath3): BrowserFilePath = EdgePath3
        End Select

        
    Case Chrome
        Select Case True
            Case fso.FileExists(ChromePath1): BrowserFilePath = ChromePath1
            Case fso.FileExists(ChromePath2): BrowserFilePath = ChromePath2
            Case fso.FileExists(ChromePath3): BrowserFilePath = ChromePath3
        End Select
    End Select
    
    BrowserVersion = fso.GetFileVersion(BrowserFilePath)
End Property

'// 出力例　"94"
Public Property Get ToMajor(Version As String)
    Dim Vers
    Vers = Split(Version, ".")
    ToMajor = Vers(0)
End Property
'// 出力例　"94.0"
Public Property Get ToMinor(Version As String)
    Dim Vers
    Vers = Split(Version, ".")
    ToMinor = Join(Array(Vers(0), Vers(1)), ".")
End Property
'// 出力例　"94.0.992"
Public Property Get ToBuild(Version As String)
    Dim Vers
    Vers = Split(Version, ".")
    ToBuild = Join(Array(Vers(0), Vers(1), Vers(2)), ".")
End Property


'// OSが64Bitかどうかを判定する
Public Property Get Is64BitOS() As Boolean
    Dim Arch As String
    '戻り値 "AMD64","IA64","x86"のいずれか
    Arch = wsh.Environment("Process").Item("PROCESSOR_ARCHITECTURE")
    '64bitOSで32bitOfficeを実行している場合、PROCESSOR_ARCHITEW6432に本来のOSのbit数が退避されているので確認
    If InStr(Arch, "64") = 0 Then Arch = wsh.Environment("Process").Item("PROCESSOR_ARCHITEW6432")
    Is64BitOS = InStr(Arch, "64")
End Property


Public Function DownloadWebDriver(Browser As BrowserName, Version As String, Optional PathSaveTo As String) As String
    Dim url As String
    If PathSaveTo = "" Then PathSaveTo = ZipPath(Browser)
    
    Select Case Browser
    Case BrowserName.Chrome
        Select Case True
            Case ToMajor(Version) < 115: url = Replace("https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip", "{version}", Version)
            Case Is64BitOS:              url = Replace("https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{version}/win64/chromedriver-win64.zip", "{version}", Version)
            Case Else:                   url = Replace("https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{version}/win32/chromedriver-win32.zip", "{version}", Version)
        End Select
        
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip", "{version}", Version)
            Case Else: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win32.zip", "{version}", Version)
        End Select
    End Select
    
    Dim Ret As Long
    DeleteUrlCacheEntry url
    Ret = URLDownloadToFile(0, url, PathSaveTo, 0, 0)
    If Ret <> 0 Then Err.Raise 4001, , "ダウンロード失敗 : " & url
    DownloadWebDriver = PathSaveTo
End Function

Public Function Extract(PathFrom As String, Optional PathTo As String) As String
    
    ' hoge.zip → hoge
    If PathTo = "" Then PathTo = Left(PathFrom, Len(PathFrom) - 4)
    
    Debug.Print "zipを展開します"
    fso.CreateFolder PathTo
    Debug.Print "    一時フォルダ : " & PathTo
    
    'PowerShellを使って展開するとマルウェア判定されたので，
    'MS非推奨だがShell.Applicationを使ってzipを解凍する
    
    On Error GoTo Catch
    'zipファイルに入っているファイルを指定したフォルダーにコピーする
    '文字列を一度()で評価してからNamespaceに渡さないとエラーが出る
    shell.Namespace((PathTo)).CopyHere shell.Namespace((PathFrom)).Items
    Extract = PathTo
    Exit Function
Catch:
    fso.DeleteFolder PathTo, True
    Err.Raise 4002, , "Zipの展開に失敗しました。原因：" & Err.Description
    Exit Function
End Function

Public Function FindExe(FolderPath) As String
    Dim f
    For Each f In fso.GetFolder(FolderPath).Files
        If f.Name Like "*.exe" Then FindExe = f.path
        If FindExe <> "" Then Exit Function
    Next

    For Each f In fso.GetFolder(FolderPath).SubFolders
        FindExe = FindExe(f)
        If FindExe <> "" Then Exit Function
    Next
End Function


Public Function RequestWebDriverVersion(ChromeVer As String) As String
    Dim http 'As XMLHTTP60
    Dim url As String
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ChromeVer
    http.Open "GET", url, False
    http.send
    
    If http.statusText = "OK" Then
        RequestWebDriverVersion = http.responseText
        Exit Function
    End If
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build.json"
    http.Open "GET", url, False
    http.send
    
    If http.statusText <> "OK" Then
        Err.Raise 4003, , "適合ドライバーの情報を取得できませんでした"
        Exit Function
    End If
    
    RequestWebDriverVersion = ParseJson(http.responseText)("builds")(ChromeVer)("version")
End Function



Public Sub InstallWebDriver(Browser As BrowserName, DriverPathTo As String)
    
    If DriverPathTo = "" Then DriverPathTo = WebDriverPath(Browser)
    
    Debug.Print "WebDriverをインストールします......"
    
    Dim BrowserVer   As String
    Dim DriverVer As String
    BrowserVer = BrowserVersion(Browser)
    Select Case Browser
        Case BrowserName.Chrome: DriverVer = RequestWebDriverVersion(ToBuild(BrowserVer))
        Case BrowserName.Edge:   DriverVer = BrowserVer
    End Select
    
    Debug.Print "   ブラウザ          : Ver. " & BrowserVer
    Debug.Print "   適合するWebDriver : Ver. " & DriverVer
    
    Dim ZipFile As String
    ZipFile = DownloadWebDriver(Browser, DriverVer)
    
    Do Until fso.FileExists(ZipFile)
        DoEvents
    Loop
    Debug.Print "   ダウンロード完了:" & ZipFile
    
    
    If Not fso.FolderExists(fso.GetParentFolderName(DriverPathTo)) Then
        Debug.Print "   WebDriverの保存先フォルダを作成します"
        CreateFolderEx fso.GetParentFolderName(DriverPathTo)
    End If
    
    Dim ExtractedFolder As String
    ExtractedFolder = Extract(ZipFile)
    
    Dim ExePath As String
    ExePath = FindExe(ExtractedFolder)
    
    If fso.FileExists(DriverPathTo) Then fso.DeleteFile DriverPathTo, True
    fso.CopyFile ExePath, DriverPathTo, True
    
    fso.DeleteFolder ExtractedFolder
    Debug.Print "    展開 : " & DriverPathTo
    Debug.Print "WebDriverを配置しました"
    Debug.Print "インストール完了"
End Sub

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
Public Sub SafeOpen(Driver As Selenium.WebDriver, Browser As BrowserName, Optional CustomDriverPath As String)
    
    Dim DriverPath As String
    DriverPath = IIf(CustomDriverPath <> "", CustomDriverPath, WebDriverPath(Browser))
    
    '// アップデート処理
    If Not IsLatestDriver(Browser, DriverPath) Then
        Dim TmpDriver As String
        If fso.FileExists(DriverPath) Then TmpDriver = BuckupTempDriver(DriverPath)
        
        Call InstallWebDriver(Browser, DriverPath)
    End If
    
    On Error GoTo Catch
    Select Case Browser
        Case BrowserName.Chrome: Driver.Start "chrome"
        Case BrowserName.Edge: Driver.Start "edge"
    End Select
    
    If TmpDriver <> "" Then Call DeleteTempDriver(TmpDriver)
    Exit Sub
    
Catch:
    If TmpDriver <> "" Then Call RollbackDriver(TmpDriver, DriverPath)
    Err.Raise Err.Number, , Err.Description
    
End Sub


'// ドライバーのバージョンを調べる
Function DriverVersion(DriverPath As String) As String
    
    If Not fso.FileExists(DriverPath) Then DriverVersion = "": Exit Function
    
    Dim TempFile
    Dim VersionInfo
    TempFile = Environ$("TMP") & "\DriverVersion_" & Format$(Now, "YYYYMMDDHHMMSS") & ".txt"
    CreateObject("WScript.Shell").Run "cmd /c " & DriverPath & " -version >" & TempFile, 0, True
    
    With fso.OpenTextFile(TempFile)
        VersionInfo = .ReadLine
        .Close
    End With
    
    fso.DeleteFile TempFile, True
    
    'バージョン情報が取得できない古いバージョンがある
    If VersionInfo = "" Then DriverVersion = "": Exit Function
    
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "\d+\.\d+\.\d+(\.\d+|)"
    
    On Error Resume Next
    DriverVersion = reg.Execute(VersionInfo)(0).Value
End Function

'// 最新のドライバーがインストールされているか調べる
Function IsLatestDriver(Browser As BrowserName, DriverPath As String) As Boolean
    Select Case Browser
    Case BrowserName.Edge
        IsLatestDriver = BrowserVersion(Edge) = DriverVersion(DriverPath)
    
    '// Chromeは末尾のバージョンがブラウザとドライバーで異なることがある
    Case BrowserName.Chrome
        IsLatestDriver = RequestWebDriverVersion(ToBuild(BrowserVersion(Chrome))) = DriverVersion(DriverPath)
    
    End Select
End Function

'// WebDriverを一時フォルダに退避させる
Function BuckupTempDriver(DriverPath As String) As String
    Dim TempFolder As String
    TempFolder = fso.BuildPath(fso.GetParentFolderName(DriverPath), fso.GetTempName)
    fso.CreateFolder TempFolder
    
    Dim TempDriver As String
    TempDriver = fso.BuildPath(TempFolder, fso.GetFileName(DriverPath))
    fso.MoveFile DriverPath, TempDriver
    
    BuckupTempDriver = TempDriver
End Function

'// 一時的に取っておいた古いWebDriverを一時フォルダからWebDriver置き場に戻す
Sub RollbackDriver(TempDriverPath As String, DriverPath As String)
    fso.CopyFile TempDriverPath, DriverPath, True
    fso.DeleteFolder fso.GetParentFolderName(TempDriverPath)
End Sub

'// 一時的に取っておいた古いWebDriverを削除する
Sub DeleteTempDriver(TempDriverPath As String)
    fso.DeleteFolder fso.GetParentFolderName(TempDriverPath)
End Sub

'簡易的なJsonパーサー
Function ParseJson(Json As String) As Object
    Dim i As Long
    i = 1
    SkipNull Json, i
    Select Case Mid(Json, i, 1)
    Case "{"
        i = i + 1
        Set ParseJson = ParseObject(Json, i)
    Case Else
        Err.Raise 4000, , "Jsonのパースに失敗"
    End Select
End Function

Private Sub SkipNull(Json, ByRef i)
    Do
        Select Case Mid(Json, i, 1)
        Case " ", vbCr, vbLf, vbTab
            i = i + 1
        Case Else
            Exit Sub
        End Select
    Loop
End Sub

Private Function ParseObject(Json As String, ByRef i) As Object
    Dim Obj As Object
    Set Obj = CreateObject("Scripting.Dictionary")
    Dim Key
    
    Do
        SkipNull Json, i
        If Mid(Json, i, 1) <> """" Then Err.Raise 4000, , "Jsonのパースに失敗"
        i = i + 1
        Key = ParseString(Json, i)
        
        SkipNull Json, i
        If Mid(Json, i, 1) <> ":" Then Err.Raise 4000, , "Jsonのパースに失敗"
        i = i + 1
        
        SkipNull Json, i
        Select Case Mid(Json, i, 1)
        Case """"
            i = i + 1
            Obj(Key) = ParseString(Json, i)
        Case "{"
            i = i + 1
            Set Obj(Key) = ParseObject(Json, i)
        Case "["
            i = i + 1
            Obj(Key) = ParseArray(Json, i)
        End Select
        
        SkipNull Json, i
        
        Select Case Mid(Json, i, 1)
        Case ","
            i = i + 1
        Case "}"
            i = i + 1
            Set ParseObject = Obj
            Exit Function
        Case Else
            Err.Raise 4000, , "Jsonのパースに失敗"
        End Select
    Loop
End Function

Private Function ParseArray(Json As String, ByRef i) As Variant
    Dim Arr
    Arr = Array()
    
     Do
        SkipNull Json, i
        Select Case Mid(Json, i, 1)
        Case """"
            i = i + 1
            ReDim Preserve Arr(0 To UBound(Arr) + 1)
            Arr(UBound(Arr)) = ParseString(Json, i)
        Case "{"
            i = i + 1
            ReDim Preserve Arr(0 To UBound(Arr) + 1)
            Set Arr(UBound(Arr)) = ParseObject(Json, i)
        Case "["
            i = i + 1
            ReDim Preserve Arr(0 To UBound(Arr) + 1)
            Arr(UBound(Arr)) = ParseArray(Json, i)
        End Select
        
        SkipNull Json, i
        
        Select Case Mid(Json, i, 1)
        Case ","
            i = i + 1
        Case "]"
            i = i + 1
            ParseArray = Arr
            Exit Function
        Case Else
            Err.Raise 4000, , "Jsonのパースに失敗"
        End Select
    Loop
End Function

Private Function ParseString(Json, i) As String
    Dim s As String
    ParseString = ""
    Do
        s = Mid(Json, i, 1)
        If s = """" Then
            i = i + 1
            Exit Function
        End If
        ParseString = ParseString & s
        i = i + 1
    Loop
End Function
