Option Compare Database
Option Explicit

Public Function Access_Class_Module_Exists(strClassName As String) As Boolean

    ' Prüft, ob ein Klassenmodul mit dem angegebenen Namen im aktuellen VBA-Projekt existiert

    Dim objComponent As Object

    On Error GoTo Fehler

    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        If objComponent.Type = vbext_ct_ClassModule Then
            If objComponent.Name = strClassName Then
                Access_Class_Module_Exists = True
                Exit Function
            End If
        End If
    Next objComponent

    Access_Class_Module_Exists = False
    Exit Function

Fehler:
    Access_Class_Module_Exists = False

End Function