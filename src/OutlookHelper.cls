Attribute VB_Name = "OutlookHelper"
Option Explicit

Private pName As String

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal value As String)
    pName = value
End Property

Public Sub Greet()
    MsgBox "Привет, " & pName & "!", vbInformation
End Sub
