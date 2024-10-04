Attribute VB_Name = "WebDriverManager4TinySelenium"
Option Explicit

Enum BrowserName
    Chrome
    Edge
End Enum

#Const DEV = 0

Private Declare PtrSafe Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" ( _
    ByVal lpszUrlName As String) As Long
    
Private Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    ByRef phReadPipe As LongPtr, _
    ByRef phWritePipe As LongPtr, _
    ByRef lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long
    
Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Any, _
    ByVal lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByRef lpEnvironment As Any, _
    ByVal lpCurrentDriectory As String, _
    ByRef lpSTARTUPINFO As STARTUPINFO, _
    ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
    
Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long
    
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As LongPtr, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    ByRef lpBytesRead As Long, _
    ByRef lpTotalBytesAvail As Long, _
    ByRef lpBytesLeftThisMessage As Long) As Long
    
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long
    
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type
    
Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0

Private Const IsSuccess = 0
Private Const Stdout = 1

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
    Dim MyDocuments As String
    MyDocuments = wsh.SpecialFolders("MyDocuments")
    Select Case Browser
        Case BrowserName.Chrome: WebDriverPath = MyDocuments & "\WebDriver\chromedriver.exe"
        Case BrowserName.Edge:   WebDriverPath = MyDocuments & "\WebDriver\edgedriver.exe"
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

'// �o�͗�@"94"
Public Property Get ToMajor(Version As String)
    Dim Vers
    Vers = Split(Version, ".")
    ToMajor = Vers(0)
End Property
'// �o�͗�@"94.0"
Public Property Get ToMinor(Version As String)
    Dim Vers
    Vers = Split(Version, ".")
    ToMinor = Join(Array(Vers(0), Vers(1)), ".")
End Property
'// �o�͗�@"94.0.992"
Public Property Get ToBuild(Version As String)
    Dim Vers
    Vers = Split(Version, ".")
    ToBuild = Join(Array(Vers(0), Vers(1), Vers(2)), ".")
End Property


'// OS��64Bit���ǂ����𔻒肷��
Public Property Get Is64BitOS() As Boolean
    Dim Arch As String
    '�߂�l "AMD64","IA64","x86"�̂����ꂩ
    Arch = wsh.Environment("Process").Item("PROCESSOR_ARCHITECTURE")
    '64bitOS��32bitOffice�����s���Ă���ꍇ�APROCESSOR_ARCHITEW6432�ɖ{����OS��bit�����ޔ�����Ă���̂Ŋm�F
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
            Case Is64BitOS:              url = Replace("https://storage.googleapis.com/chrome-for-testing-public/{version}/win64/chromedriver-win64.zip", "{version}", Version)
            Case Else:                   url = Replace("https://storage.googleapis.com/chrome-for-testing-public/{version}/win32/chromedriver-win32.zip", "{version}", Version)
        End Select
        
    Case BrowserName.Edge
        Select Case Is64BitOS
            Case True: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip", "{version}", Version)
            Case Else: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win32.zip", "{version}", Version)
        End Select
    End Select
    
    If fso.FolderExists(PathSaveTo) Then fso.DeleteFolder PathSaveTo, True
    DeleteUrlCacheEntry url
    
    Dim http
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
    
    If http.Status <> 200 Then
        Err.Raise 4001, , "�_�E�����[�h���s : " & url
        Exit Function
    End If
    
    Const adTypeBinary = 1
    Const adSaveCreateNotExist = 2
    With CreateObject("ADODB.Stream")
        .Type = adTypeBinary
        .Open
        .Position = 0
        .Write http.responseBody
        .SaveToFile PathSaveTo, adSaveCreateNotExist
        .Close
    End With

    DownloadWebDriver = PathSaveTo
End Function

Public Function Extract(PathFrom As String, Optional PathTo As String) As String
    
    ' hoge.zip �� hoge
    If PathTo = "" Then PathTo = Left(PathFrom, Len(PathFrom) - 4)
    
    Debug.Print "zip��W�J���܂�"
    If fso.FolderExists(PathTo) Then fso.DeleteFolder PathTo, True
    fso.CreateFolder PathTo
    Debug.Print "    �ꎞ�t�H���_ : " & PathTo

    Do Until fso.FolderExists(PathTo)
        Sleep 100
        DoEvents
    Loop
    
    'PowerShell���g���ēW�J����ƃ}���E�F�A���肳�ꂽ�̂ŁC
    'MS�񐄏�����Shell.Application���g����zip���𓀂���
    
    On Error GoTo Catch
    'zip�t�@�C���ɓ����Ă���t�@�C�����w�肵���t�H���_�[�ɃR�s�[����
    '���������x()�ŕ]�����Ă���Namespace�ɓn���Ȃ��ƃG���[���o��
    shell.Namespace((PathTo)).CopyHere shell.Namespace((PathFrom)).Items
    Extract = PathTo
    Exit Function
Catch:
    fso.DeleteFolder PathTo, True
    Err.Raise 4002, , "Zip�̓W�J�Ɏ��s���܂����B�����F" & Err.Description
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
    Dim http
    Dim url As String

    ChromeVer = ToBuild(ChromeVer)
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    'ChromeVer.114�܂ł�WebDriver�z�z�T�C�g
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ChromeVer
    http.Open "GET", url, False
    http.send
    
    If http.Status = 200 Then
        RequestWebDriverVersion = http.responseText
        Exit Function
    End If
    
    'ChromeVer.115�����WebDriver�z�z�T�C�g
    url = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build.json"
    http.Open "GET", url, False
    http.send
    
    If http.Status <> 200 Then
        Err.Raise 4003, , "�K���h���C�o�[�̏����擾�ł��܂���ł���"
        Exit Function
    End If

    RequestWebDriverVersion = ParseJson(http.responseText)("builds")(ChromeVer)("version")
End Function



Public Sub InstallWebDriver(Browser As BrowserName, Optional DriverPathTo As String)
    
    If DriverPathTo = "" Then DriverPathTo = WebDriverPath(Browser)
    
    Debug.Print "WebDriver���C���X�g�[�����܂�......"
    
    Dim BrowserVer   As String
    Dim DriverVer As String
    BrowserVer = BrowserVersion(Browser)
    Select Case Browser
        Case BrowserName.Chrome: DriverVer = RequestWebDriverVersion(BrowserVer)
        Case BrowserName.Edge:   DriverVer = BrowserVer
    End Select
    
    Debug.Print "   �u���E�U          : Ver. " & BrowserVer
    Debug.Print "   �K������WebDriver : Ver. " & DriverVer
    
    Dim ZipFile As String
    ZipFile = DownloadWebDriver(Browser, DriverVer)
    
    Do Until fso.FileExists(ZipFile)
        DoEvents
    Loop
    Debug.Print "   �_�E�����[�h����:" & ZipFile
    
    
    If Not fso.FolderExists(fso.GetParentFolderName(DriverPathTo)) Then
        Debug.Print "   WebDriver�̕ۑ���t�H���_���쐬���܂�"
        CreateFolderEx fso.GetParentFolderName(DriverPathTo)
        Do Until fso.FolderExists(fso.GetParentFolderName(DriverPathTo))
            Sleep 100
            DoEvents
        Loop
    End If
    
    Dim ExtractedFolder As String
    ExtractedFolder = Extract(ZipFile)
    
    Dim ExePath As String
    ExePath = FindExe(ExtractedFolder)
    
    If fso.FileExists(DriverPathTo) Then fso.DeleteFile DriverPathTo, True
    fso.CopyFile ExePath, DriverPathTo, True
    
    fso.DeleteFolder ExtractedFolder
    Debug.Print "    �W�J : " & DriverPathTo
    Debug.Print "WebDriver��z�u���܂���"
    Debug.Print "�C���X�g�[������"
End Sub

Public Sub CreateFolderEx(path_folder As String)
    '// �e�t�H���_���k��Ȃ��Ȃ�Ƃ���܂ōċA�ŒH��
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    '// �r���̑��݂��Ȃ��t�H���_���쐬���Ȃ���~��Ă���
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub



Public Sub SafeOpen(Driver As WebDriver, Browser As BrowserName, Optional CustomDriverPath As String, Optional CapabilityArgs As String)
    On Error GoTo Catch
    Dim driverPath As String
    driverPath = IIf(CustomDriverPath <> "", CustomDriverPath, WebDriverPath(Browser))
    
    '// �A�b�v�f�[�g����
    If Not IsLatestDriver(Browser, driverPath) Then
        Dim TmpDriver As String
        If fso.FileExists(driverPath) Then TmpDriver = BuckupTempDriver(driverPath)
        
        Call InstallWebDriver(Browser, driverPath)
    End If
    
    Select Case Browser
        Case BrowserName.Chrome: Driver.Chrome driverPath
        Case BrowserName.Edge:   Driver.Edge driverPath
    End Select

    If CapabilityArgs <> "" Then
        Dim cap As Capabilities
        Set cap = Driver.CreateCapabilities()
        cap.SetArguments CapabilityArgs
        Driver.OpenBrowser cap
    Else
        Driver.OpenBrowser
    End If
    
    If TmpDriver <> "" Then Call DeleteTempDriver(TmpDriver)
    Exit Sub
    
Catch:
    If TmpDriver <> "" Then Call RollbackDriver(TmpDriver, driverPath)
    Err.Raise Err.Number, , Err.Description
    
End Sub


'// �h���C�o�[�̃o�[�W�����𒲂ׂ�
Function DriverVersion(driverPath As String) As String
    
    If Not fso.FileExists(driverPath) Then DriverVersion = "": Exit Function
    
    Dim Res
    Res = ReadStdOut("""" & driverPath & """ --version")
    If Res(IsSuccess) Then
        Dim reg
        Set reg = CreateObject("VBScript.RegExp")
        reg.Pattern = "\d+\.\d+\.\d+(\.\d+|)"
        
        On Error Resume Next
        DriverVersion = reg.Execute(Res(Stdout))(0).value
    Else
        DriverVersion = ""
    End If
End Function

'// �ŐV�̃h���C�o�[���C���X�g�[������Ă��邩���ׂ�
Function IsLatestDriver(Browser As BrowserName, driverPath As String) As Boolean
    Select Case Browser
    Case BrowserName.Edge
        IsLatestDriver = BrowserVersion(Edge) = DriverVersion(driverPath)
    
    '// Chrome�͖����̃o�[�W�������u���E�U�ƃh���C�o�[�ňقȂ邱�Ƃ�����
    Case BrowserName.Chrome
        IsLatestDriver = RequestWebDriverVersion(BrowserVersion(Chrome)) = DriverVersion(driverPath)
    
    End Select
End Function

'// WebDriver���ꎞ�t�H���_�ɑޔ�������
Function BuckupTempDriver(driverPath As String) As String
    Dim TempFolder As String
    TempFolder = fso.BuildPath(fso.GetParentFolderName(driverPath), fso.GetTempName)
    fso.CreateFolder TempFolder

    Do Until fso.FolderExists(TempFolder)
        Sleep 100
        DoEvents
    Loop
    
    Dim TempDriver As String
    TempDriver = fso.BuildPath(TempFolder, fso.GetFileName(driverPath))
    fso.MoveFile driverPath, TempDriver
    
    BuckupTempDriver = TempDriver
End Function

'// �ꎞ�I�Ɏ���Ă������Â�WebDriver���ꎞ�t�H���_����WebDriver�u����ɖ߂�
Sub RollbackDriver(TempDriverPath As String, driverPath As String)
    fso.CopyFile TempDriverPath, driverPath, True
    fso.DeleteFolder fso.GetParentFolderName(TempDriverPath)
End Sub

'// �ꎞ�I�Ɏ���Ă������Â�WebDriver���폜����
Sub DeleteTempDriver(TempDriverPath As String)
    fso.DeleteFolder fso.GetParentFolderName(TempDriverPath)
End Sub

'�ȈՓI��Json�p�[�T�[
Function ParseJson(Json As String) As Object
    Dim i As Long
    Dim s0 As String
    Dim s1 As String
    i = 1
    Do While i <= Len(Json)
        SkipNull Json, i
        Select Case Mid(Json, i, 1)
        Case "{"
            i = i + 1
            Set ParseJson = ParseObject(Json, i)
            Exit Function
        End Select
    Loop
    
End Function

Private Sub SkipNull(Json, ByRef i)
    Dim s As String
    s = Mid(Json, i, 1)
    Do While s = " " Or s = vbCr Or s = vbLf Or s = vbTab
        i = i + 1
        s = Mid(Json, i, 1)
    Loop
    
End Sub

Private Function ParseObject(Json As String, ByRef i)
    Dim Obj As Object
    Set Obj = CreateObject("Scripting.Dictionary")
    Dim key
    
    Do
        SkipNull Json, i
        If Mid(Json, i, 1) <> """" Then Err.Raise 4000, , "Json�̃p�[�X�Ɏ��s"
        i = i + 1
        key = ParseString(Json, i)
        
        SkipNull Json, i
        If Mid(Json, i, 1) <> ":" Then Err.Raise 4000, , "Json�̃p�[�X�Ɏ��s"
        i = i + 1
        
        SkipNull Json, i
        Select Case Mid(Json, i, 1)
        Case """"
            i = i + 1
            Let Obj(key) = ParseString(Json, i)
        Case "{"
            i = i + 1
            Set Obj(key) = ParseObject(Json, i)
        Case "["
            i = i + 1
            Set Obj(key) = ParseArray(Json, i)
        End Select
        
        SkipNull Json, i
        
        Select Case Mid(Json, i, 1)
        Case ","
            i = i + 1
        Case "}"
            i = i + 1
            Set ParseObject = Obj
            Exit Do
        End Select
    Loop
End Function

Private Function ParseArray(Json As String, ByRef i)
    Dim Arr As Collection
    Set Arr = New Collection
    
     Do
        SkipNull Json, i
        Select Case Mid(Json, i, 1)
        Case """"
            i = i + 1
            Arr.Add ParseString(Json, i)
        Case "{"
            i = i + 1
            Arr.Add ParseObject(Json, i)
        Case "["
            i = i + 1
            Arr.Add ParseArray(Json, i)
        End Select
        
        SkipNull Json, i
        
        Select Case Mid(Json, i, 1)
        Case ","
            i = i + 1
        Case "]"
            i = i + 1
            Set ParseArray = Arr
            Exit Do
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
            Exit Do
        End If
        ParseString = ParseString & s
        i = i + 1
    Loop
End Function

'�R�}���h�����s�������̕W���o�͂�ǂݎ��֐�
'�߂�l Array(�����������ǂ���,�W���o��)
Function ReadStdOut(cmd As String)
    Const FAILED = 0
    Dim Result_IsSuccess As Boolean
    Dim Result_StdOut    As String
    
    Dim ReadPipe  As LongPtr
    Dim WritePipe As LongPtr
    Dim sa As SECURITY_ATTRIBUTES
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    sa.lpSecurityDescriptor = 0
    
    If CreatePipe(ReadPipe, WritePipe, sa, 0) = FAILED Then
        GoTo finally
    End If
    
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    si.hStdInput = ReadPipe
    si.hStdOutput = WritePipe
    si.hStdError = WritePipe
    si.wShowWindow = SW_HIDE
    
    cmd = "/c " & cmd
    If CreateProcess("C:\Windows\System32\cmd.exe", cmd, 0&, 0&, 1&, 0&, 0&, "C:\", si, pi) = FAILED Then
        GoTo finally
    End If
    
    CloseHandle pi.hThread
    pi.hThread = 0
    
    If WaitForSingleObject(pi.hProcess, 1000) <> 0 Then
        GoTo finally
    End If
    
    Dim ReadBuf() As Byte
    Dim TotalLength As Long
    Dim Length As Long
    If PeekNamedPipe(ReadPipe, 0, 0, 0, TotalLength, 0) = FAILED Then
        GoTo finally
    End If
    If 0 < TotalLength Then
        ReDim ReadBuf(0 To TotalLength - 1) As Byte
        If ReadFile(ReadPipe, ReadBuf(0), UBound(ReadBuf), 0&, 0&) = FAILED Then
            GoTo finally
        End If
    End If
    
    Result_IsSuccess = True
    Result_StdOut = StrConv(ReadBuf, vbUnicode)

finally:
    If WritePipe <> 0 Then CloseHandle WritePipe
    If ReadPipe <> 0 Then CloseHandle ReadPipe
    If pi.hProcess <> 0 Then CloseHandle pi.hProcess
    
    ReadStdOut = Array(Result_IsSuccess, Result_StdOut)
End Function
