Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Database

Const c_strForm_Log As String = "frm_Log"

Dim Writing As New clsWriting

' Konstruktor
Private Sub Class_Initialize()
    ' Prüft, ob das Formular geöffnet ist, und öffnet es falls notwendig
    If Not IsFormOpen(c_strForm_Log) Then
        DoCmd.OpenForm c_strForm_Log
    End If
End Sub

' Prüft, ob ein Formular geöffnet ist
Private Function IsFormOpen(strFormName As String) As Boolean
    IsFormOpen = (SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0)
End Function

' Eigenschaften
Public Property Get FullText() As String
    ' Gibt den vollständigen Log-Text zurück
    FullText = Forms(c_strForm_Log).txtLog.value
End Property

Public Property Get LastLine() As String
    ' Gibt die letzte Zeile des Logs zurück
    Dim arrLines As Variant
    arrLines = Split(Forms(c_strForm_Log).txtLog.value, vbCrLf)
    If UBound(arrLines) >= 0 Then
        LastLine = arrLines(UBound(arrLines))
    Else
        LastLine = ""
    End If
End Property

' Methoden
Public Sub WriteText(strText As String)
    ' Fügt Text ohne Zeilenumbruch hinzu
    Forms(c_strForm_Log).txtLog.value = Forms(c_strForm_Log).txtLog.value & strText
End Sub

Public Sub WriteLine(strText As String)
    ' Fügt eine Zeile mit Zeilenumbruch hinzu
    Forms(c_strForm_Log).txtLog.value = Forms(c_strForm_Log).txtLog.value & Get_WrapLine(strText) & vbCrLf
End Sub

Public Sub WriteEmptyLine()
    ' Fügt eine leere Zeile hinzu
    Forms(c_strForm_Log).txtLog.value = Forms(c_strForm_Log).txtLog.value & vbCrLf
End Sub
Private Function Get_WrapLine(strLineText As String) As String
 
    Get_WrapLine = Writing.Get_Timestamp() & " " & Writing.Get_Log_Prefix_ByType(ptInfo) & " " & strLineText

End Function