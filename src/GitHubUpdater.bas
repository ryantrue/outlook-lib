Attribute VB_Name = "GitHubUpdater"
Option Explicit

Private Const GITHUB_RAW_URL As String = "https://raw.githubusercontent.com/ryantrue/outlook-lib/main/"
Private Const VERSION_URL As String = GITHUB_RAW_URL & "VERSION.txt"
Private Const COMPONENTS_TO_UPDATE As String = "OutlookCoreModule.bas,OutlookHelper.cls"
Private Const VERSION_FILE As String = "C:\Temp\outlook-lib-version.txt"
Private Const TEMP_DIR As String = "C:\Temp\outlook-lib\"

Public Sub AutoUpdateLibrary()
    On Error GoTo HandleError

    EnsureTempFolderExists

    Dim localVersion As String: localVersion = ReadLocalVersion()
    Dim remoteVersion As String: remoteVersion = DownloadTextFile(VERSION_URL)

    If Trim(localVersion) <> Trim(remoteVersion) Then
        Dim file As Variant
        For Each file In Split(COMPONENTS_TO_UPDATE, ",")
            Dim fileName As String: fileName = Trim(file)
            Dim downloadURL As String: downloadURL = GITHUB_RAW_URL & fileName
            Dim localPath As String: localPath = TEMP_DIR & fileName

            DownloadFile downloadURL, localPath
            ReplaceComponentFromFile localPath
        Next

        SaveLocalVersion remoteVersion
        MsgBox "Библиотека outlook-lib обновлена до версии " & remoteVersion, vbInformation
    End If
    Exit Sub

HandleError:
    MsgBox "Ошибка обновления outlook-lib: " & Err.Description, vbCritical
End Sub

Private Sub EnsureTempFolderExists()
    If Dir(TEMP_DIR, vbDirectory) = "" Then MkDir TEMP_DIR
End Sub

Private Function ReadLocalVersion() As String
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open VERSION_FILE For Input As #f
    Line Input #f, ReadLocalVersion
    Close #f
End Function

Private Sub SaveLocalVersion(version As String)
    Dim f As Integer: f = FreeFile
    Open VERSION_FILE For Output As #f
    Print #f, version
    Close #f
End Sub

Private Function DownloadTextFile(url As String) As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False
        .Send
        If .Status = 200 Then
            DownloadTextFile = .responseText
        Else
            Err.Raise vbObjectError + 100, , "HTTP " & .Status & " при загрузке версии"
        End If
    End With
End Function

Private Sub DownloadFile(url As String, localPath As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")

    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False
        .Send
        If .Status <> 200 Then Err.Raise vbObjectError + 101, , "Не удалось загрузить: " & url

        stream.Type = 1
        stream.Open
        stream.Write .responseBody
        stream.SaveToFile localPath, 2
        stream.Close
    End With
End Sub

Private Sub ReplaceComponentFromFile(filePath As String)
    Dim vbProj As Object: Set vbProj = Application.VBE.ActiveVBProject
    Dim fileName As String: fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    Dim componentName As String: componentName = Replace(fileName, ".bas", "")
    componentName = Replace(componentName, ".cls", "")

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Name = componentName Then
            vbProj.VBComponents.Remove comp
            Exit For
        End If
    Next

    vbProj.VBComponents.Import filePath
End Sub
