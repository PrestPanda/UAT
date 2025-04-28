Option Compare Database
Option Explicit

Public Function Array_HasEntries(varArray As Variant) As Boolean

    ' Überprüft, ob ein übergebenes Array Einträge enthält

    On Error GoTo Fehler

    If IsArray(varArray) Then
        If Not IsEmpty(varArray) Then
            If (UBound(varArray) >= LBound(varArray)) Then
                Array_HasEntries = True
                Exit Function
            End If
        End If
    End If

    Array_HasEntries = False
    Exit Function

Fehler:
    Array_HasEntries = False

End Function