Option Compare Database
Option Explicit

Public Function Array_HasEntries(varArray As Variant) As Boolean

    ' �berpr�ft, ob ein �bergebenes Array Eintr�ge enth�lt

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