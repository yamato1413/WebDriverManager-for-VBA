Attribute VB_Name = "WebDriverManager"
Option Explicit

Enum BrowserName
    Chrome
    Edge
End Enum


'// ƒtƒ@ƒCƒ‹ƒ_ƒEƒ“ƒ[ƒh—p‚ÌWin32API
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#Else
Private Declare  Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare  Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If



Private Property Get fso() 'As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property
Private Property Get wsh() 'As WshShell
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("WScript.Shell")
    Set wsh = obj
End Property

'// ƒ_ƒEƒ“ƒ[ƒh‚µ‚½WebDriver‚Ìzip‚Ì•Û‘¶êŠ
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


'// WebDriver‚ÌÀsƒtƒ@ƒCƒ‹‚Ì•Û‘¶êŠ‚ğƒŒƒWƒXƒgƒŠ‚É‹L˜^‚µ‚Ä‚¢‚éB
'// ƒfƒtƒHƒ‹ƒg‚ÍƒhƒLƒ…ƒƒ“ƒgƒtƒHƒ‹ƒ_
'// ‚±‚ÌƒpƒX‚ğ‘‚«Š·‚¦‚éƒvƒƒV[ƒWƒƒ‚ÍˆÈ‰º‚Ì’Ê‚è
'// EProperty Let WebDriverPath
'// EInstallWebDriver
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
    
    If version = "" Then Err.Raise 4000, , "ƒo[ƒWƒ‡ƒ“î•ñ‚ªæ“¾‚Å‚«‚Ü‚¹‚ñ‚Å‚µ‚½"
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



'// ‘æ3ˆø”‚É‚ÄƒpƒX‚ğw’è‚·‚ê‚Î”CˆÓ‚ÌêŠ‚É”CˆÓ‚Ì–¼‘O‚Å•Û‘¶‚Å‚«‚éB
'// g—p—á DownloadWebDriver Edge, "94.0.992.31", "C:\Users\yamato\Desktop\edge.zip"
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
    
    If path_save_to = "" Then path_save_to = ZipPath(browser)   'ƒfƒtƒH‚Í"C:Users\USERNAME\Downloads\~~~.zip"
    
    DeleteUrlCacheEntry url
    Dim ret As Long
    ret = URLDownloadToFile(0, url, path_save_to, 0, 0)
    If ret <> 0 Then Err.Raise 4001, , "ƒ_ƒEƒ“ƒ[ƒh¸”s : " & url
    
    DownloadWebDriver = path_save_to
End Function



'// zip‚©‚ç’†g‚ğæ‚èo‚µ‚Äw’è‚ÌêŠ‚ÉÀsƒtƒ@ƒCƒ‹‚ğ“WŠJ‚·‚é
'// chromedriver.exe(ƒfƒtƒHƒ‹ƒg‚Ì–¼‘O)‚ª‚ ‚é‚Æ‚±‚ë‚Échromedriver_94.exe‚Æ‚©‚Å“WŠJ‚Å‚«‚é‚æ‚¤‚ÉA
'// Œ³‚ÌÀsƒtƒ@ƒCƒ‹‚ğã‘‚«‚µ‚È‚¢‚æ‚¤‚Éˆê“xtempƒtƒHƒ‹ƒ_‚ğì‚Á‚Ä‚©‚çÀsƒtƒ@ƒCƒ‹‚ğ–Ú“I‚ÌƒpƒX‚ÖˆÚ‚·
'// g—p—á Extract "C:\Users\yamato\Downloads\chromedriver_win32.zip","C:\Users\yamato\Downloads\chromedriver_94.exe"
Sub Extract(path_zip As String, path_save_to As String)
    Dim folder_temp
    folder_temp = fso.GetParentFolderName(path_save_to) & "\" & fso.GetTempName
    fso.CreateFolder folder_temp
    
    'Shell.Application‚ğg‚¤•û–@‚ÍMS”ñ„§‚ç‚µ‚¢‚Ì‚ÅPowerShell‚Å“WŠJ‚·‚é
    Dim command As String, ex As Object 'WshExec
    command = "Expand-Archive -Path " & path_zip & " -DestinationPath " & folder_temp & " -Force"
    Set ex = wsh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & command)
    
    '// ƒRƒ}ƒ“ƒh¸”s
    If ex.Status = WshFailed Then: Err.Raise 4002, , "Zip‚Ì“WŠJ‚É¸”s‚µ‚Ü‚µ‚½": Exit Sub
    
    Do While ex.Status = 0 'WshRunning
        DoEvents
    Loop
    
    Dim path_exe_from As String, path_exe_to As String
    path_exe_from = folder_temp & "\" & Dir(folder_temp & "\*.exe")
    
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
        Err.Raise 4003, "ƒT[ƒo[‚Ö‚ÌÚ‘±‚É¸”s‚µ‚Ü‚µ‚½"
        Exit Function
    End If

    RequestWebDriverVersion = http.responseText
End Function



Sub InstallWebDriver(browser As BrowserName, Optional path_driver As String)
    Debug.Print "WebDriver‚ğƒCƒ“ƒXƒg[ƒ‹‚µ‚Ü‚·......"
    
    Dim ver_browser   As String
    Dim ver_webdriver As String
    ver_browser = BrowserVersion(browser)
    Select Case browser
        Case BrowserName.Chrome: ver_webdriver = RequestWebDriverVersion(BrowserVersionToBuild(browser))
        Case BrowserName.Edge:   ver_webdriver = ver_browser
    End Select
    
    Debug.Print "   ƒuƒ‰ƒEƒU          : Ver. " & ver_browser
    Debug.Print "   “K‡‚·‚éWebDriver : Ver. " & ver_webdriver
    
    Dim path_zip As String
    path_zip = DownloadWebDriver(browser, ver_webdriver)
    
    Do Until fso.FileExists(ZipPath(browser))
        DoEvents
    Loop
    Debug.Print "   ƒ_ƒEƒ“ƒ[ƒhŠ®—¹:" & path_zip
    
    If path_driver <> "" Then WebDriverPath(browser) = path_driver
    
    If Not fso.FolderExists(fso.GetParentFolderName(WebDriverPath(browser))) Then
        Debug.Print "   WebDriver‚Ì•Û‘¶æƒtƒHƒ‹ƒ_‚ğì¬‚µ‚Ü‚·"
        CreateFolderEx fso.GetParentFolderName(WebDriverPath(browser))
    End If
    
    Extract path_zip, WebDriverPath(browser)
    
    Debug.Print "ƒCƒ“ƒXƒg[ƒ‹Š®—¹"
End Sub



'ƒpƒX‚ÉŠÜ‚Ü‚ê‚é‘S‚Ä‚ÌƒtƒHƒ‹ƒ_‚Ì‘¶İŠm”F‚ğ‚µ‚ÄƒtƒHƒ‹ƒ_‚ğì‚éŠÖ”
Sub CreateFolderEx(path_folder As String)
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub


Sub SafeOpen(Driver As WebDriver, browser As BrowserName, Optional path_driver As String)
    If path_driver = "" Then path_driver = WebDriverPath(browser)
    
    If Not fso.FileExists(WebDriverPath(browser)) Then
        Debug.Print "WebDriver‚ªŒ©‚Â‚©‚è‚Ü‚¹‚ñ"
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




        ~             ­               @              6                                                               