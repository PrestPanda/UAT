Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Database

Const c_strForm_Log As String = "frm_Log"

Dim Writing As New clsWriting


Private Sub Class_Initialize()

    ' Pr�ft, ob das Formular ge�ffnet ist, und �ffnet es falls notwendig
    If Not IsFormOpen(c_strForm_Log) Then
        DoCmd.OpenForm c_strForm_Log
    End If
    
End Sub

' Pr�ft, ob ein Formular ge�ffnet ist
Private Function IsFormOpen(strFormName As String) As Boolean
    IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0)
End Function

' Eigenschaften
Public Property Get FullText() As String
    ' Gibt den vollst�ndigen Log-Text zur�ck
    FullText = Forms(c_strForm_Log).txtLog.Value
End Property
Public Property Let FullText(strText As String)

    Forms(c_strForm_Log).txtLog.Value = strText

End Property

Public Property Get LastLine() As String
    ' Gibt die letzte Zeile des Logs zur�ck
    Dim arrLines As Variant
    arrLines = Split(Forms(c_strForm_Log).txtLog.Value, vbCrLf)
    If UBound(arrLines) >= 0 Then
        LastLine = arrLines(UBound(arrLines))
    Else
        LastLine = ""
    End If
End Property

' Methoden
Public Sub WriteText(strText As String)
    ' F�gt Text ohne Zeilenumbruch hinzu
    Forms(c_strForm_Log).txtLog.Value = Forms(c_strForm_Log).txtLog.Value & strText

    ' Setzt den Fokus ans Ende der Textbox
    Access_TextBox_MoveToEnd c_strForm_Log, "txtLog"

End Sub

Public Sub WriteLine(strText As String)
    ' F�gt eine Zeile mit Zeilenumbruch hinzu
    Forms(c_strForm_Log).txtLog.Value = Forms(c_strForm_Log).txtLog.Value & Get_WrapLine(strText) & vbCrLf

    ' Setzt den Fokus ans Ende der Textbox
    Access_TextBox_MoveToEnd c_strForm_Log, "txtLog"

End Sub

Public Sub WriteEmptyLine()
    ' F�gt eine leere Zeile hinzu
    Forms(c_strForm_Log).txtLog.Value = Forms(c_strForm_Log).txtLog.Value & vbCrLf

    ' Setzt den Fokus ans Ende der Textbox
    Access_TextBox_MoveToEnd c_strForm_Log, "txtLog"

End Sub


Public Sub Write_Form_Opened(strFormName As String)

    WriteLine Writing.Get_Message_Form_Opened(strFormName)

End Sub

Private Function Get_WrapLine(strLineText As String) As String
 
    Get_WrapLine = Writing.Get_Timestamp() & " " & Writing.Get_Log_Prefix_ByType(ptInfo) & " " & strLineText

End Function
Public Sub Write_ProgressBar(strTitle As String, lngCurrentItem As Long, lngMaxItems As Long)
    
    WriteLine strTitle
    WriteLine Writing.Get_ProgressBar(lngCurrentItem, lngMaxItems, pboLog, pbdesDot)
    
End Sub
Public Sub Update_ProgressBar(strTitle As String, lngCurrentItem As Long, lngMaxItems As Long)
     
    FullText = Replace(FullText, LastLine, Writing.Get_ProgressBar(lngCurrentItem, lngMaxItems, pboLog, pbdesDot))

End Sub