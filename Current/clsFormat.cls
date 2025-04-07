Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function Format_AddTabToLines(ByVal strInputText As String) As String
    ' Fügt jeder Zeile des Eingabetextes ein Tabulatorzeichen (vbTab) voran
    
    Dim arrLines() As String
    Dim strOutputText As String
    Dim intIndex As Integer
    
    arrLines = Split(strInputText, vbCrLf)
    
    For intIndex = LBound(arrLines) To UBound(arrLines)
        arrLines(intIndex) = vbTab & arrLines(intIndex)
    Next intIndex
    
    strOutputText = Join(arrLines, vbCrLf)
    
    Format_AddTabToLines = strOutputText
    
End Function