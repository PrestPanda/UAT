Option Compare Database
Option Explicit

Public Function Access_Table_Exists(strTableName As String) As Boolean

    ' Prüft, ob eine Tabelle mit dem angegebenen Namen in der aktuellen Datenbank existiert

    Dim objTableDef As DAO.TableDef

    On Error GoTo Fehler

    For Each objTableDef In CurrentDb.TableDefs
        If objTableDef.Name = strTableName Then
            Access_Table_Exists = True
            Exit Function
        End If
    Next objTableDef

    Access_Table_Exists = False
    Exit Function

Fehler:
    Access_Table_Exists = False

End Function