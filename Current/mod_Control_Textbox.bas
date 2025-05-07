Option Compare Database
Option Explicit

Public Sub Access_TextBox_MoveToEnd(strFormName As String, strControlName As String)

    ' Setzt den Fokus in die Textbox und bewegt den Cursor ans Ende des Textes

    With Forms(strFormName).Controls(strControlName)
        .SetFocus
        DoEvents ' Warten, bis Access den Fokus verarbeitet hat

        ' Prüfen, ob die Text-Eigenschaft zugreifbar ist
        Dim lngTextLength As Long
        On Error Resume Next
        lngTextLength = Len(.Text)
        If Err.Number <> 0 Then
            ' Fallback: Value verwenden
            Err.Clear
            lngTextLength = Len(.Value)
        End If
        On Error GoTo 0

        .SelStart = lngTextLength
    End With

End Sub