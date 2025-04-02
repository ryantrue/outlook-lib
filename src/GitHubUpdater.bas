Attribute VB_Name = "GitHubUpdater"
Option Explicit

' Требуется: Microsoft Visual Basic for Applications Extensibility 5.3
' И включённый доступ в Trust Center к объектной модели

Private Const GITHUB_RAW_URL As String = "https://raw.githubusercontent.com/ryantrue/outlook-lib/main/src/"
Private Const VERSION_URL As String = "https://raw.githubusercontent.com/ryantrue/outlook-lib/main/VERSION.txt"
Private Const MODULES_TO_UPDATE As String = "OutlookCoreModule.bas,OutlookHelper.cls"

Private Const LOCAL_ROOT_DIR As String = "C:\Temp\outlook-lib\"
Private Const LOCAL_VERSION_FILE As String = LOCAL_ROOT_DIR & "VERSION.txt"

Public Sub AutoUpdateLibrary()
    On Error GoTo HandleError

    EnsureDirectoryExists LOCAL_ROOT_DIR
    EnsureVersionFileExists LOCAL_VERSION_FILE

    Dim localVersion As String: localVersion = ReadTextFile(LOCAL_VERSION_FILE)
    Dim remoteVersion As String: remoteVersion = DownloadTextFile(VERSION_URL)

    If Trim(localVersion) <> Trim(remoteVersion) Then
        Dim file As Variant
        For Each file In Split(MODULES_TO_UPDATE, ",")
            Dim fileName As String: fileName = Trim(file)
            Dim remoteURL As String: remoteURL = GITHUB_RAW_URL & fileName
            Dim localPath As String: localPath = LOCAL_ROOT_DIR & fileName

            DownloadFile remoteURL, localPath
            ReplaceVbaComponent localPath
        Next

        SaveTextFile LOCAL_VERSION_FILE, remoteVersion
        MsgBox "Библиотека outlook-lib обновлена до версии " & remoteVersion, vbInformation
    End If
    Exit Sub

HandleError:
    MsgBox "Ошибка обновления библиотеки: " & Err.Description, vbCritical
End Sub

Private Sub EnsureDirectoryExists(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
End Sub

Private Sub EnsureVersionFileExists(filePath As String)
    If Dir(filePath) = "" Then SaveTextFile filePath, "0.0.0"
End Sub

Private Function ReadTextFile(filePath As String) As String
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open filePath For Input As #f
    Line Input #f, ReadTextFile
    Close #f
End Function

Private Sub SaveTextFile(filePath As String, content As String)
    Dim f As Integer: f = FreeFile
    Open filePath For Output As #f
    Print #f, content
    Close #f
End Sub

Private Function DownloadTextFile(url As String) As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False
        .Send
        If .Status = 200 Then
            DownloadTextFile = .responseText
        Else
            Err.Raise vbObjectError + 100, , "Ошибка загрузки (HTTP " & .Status & ") → " & url
        End If
    End With
End Function

Private Sub DownloadFile(url As String, localPath As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False
        .Send
        If .Status <> 200 Then Err.Raise vbObjectError + 101, , "Ошибка загрузки файла: " & url

        stream.Type = 1
        stream.Open
        stream.Write .responseBody
        stream.SaveToFile localPath, 2
        stream.Close
    End With
End Sub

Private Sub ReplaceVbaComponent(filePath As String)
    Dim vbProj As Object: Set vbProj = Application.VBE.ActiveVBProject
    Dim fileName As String: fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    Dim componentName As String: componentName = Replace(Replace(fileName, ".bas", ""), ".cls", "")

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Name = componentName Then
            vbProj.VBComponents.Remove comp
            Exit For
        End If
    Next

    vbProj.VBComponents.Import filePath
End Sub
