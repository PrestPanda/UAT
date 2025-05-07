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

Public Sub Access_Relation_Create( _
    ByVal strChildTable As String, _
    ByVal strChildField As String, _
    ByVal strParentTable As String, _
    ByVal strParentField As String, _
    Optional ByVal blnEnforce As Boolean = True)
    ' Erstellt eine Fremdschlüsselbeziehung zwischen zwei Tabellen

    Dim objDB As DAO.Database
    Dim objRel As DAO.Relation
    Dim objFld As DAO.Field

    Set objDB = CurrentDb

    ' Vorhandene Relation mit demselben Namen löschen
    On Error Resume Next
    objDB.Relations.Delete strChildTable & "_" & strChildField & "_FK"
    On Error GoTo 0

    ' Neue Relation anlegen mit Cascade-Regeln
    Set objRel = objDB.CreateRelation( _
        strChildTable & "_" & strChildField & "_FK", _
        strChildTable, _
        strParentTable, _
        dbRelationUpdateCascade Or dbRelationDeleteCascade)

    ' Fremdschlüsselfeld definieren
    Set objFld = objRel.CreateField(strChildField)
    objFld.ForeignName = strParentField
    objRel.Fields.Append objFld

    ' Durchsetzung der referenziellen Integrität
    If Not blnEnforce Then
        objRel.Attributes = objRel.Attributes Or dbRelationDontEnforce
    End If

    objDB.Relations.Append objRel

    ' Objekte freigeben
    Set objFld = Nothing
    Set objRel = Nothing
    Set objDB = Nothing
End Sub

Public Sub Access_LookupField_Create( _
    ByVal strChildTable As String, _
    ByVal strChildField As String, _
    ByVal strLookupTable As String, _
    ByVal strLookupKey As String, _
    ByVal strLookupDisplay As String)
    ' Richtet ein Nachschlagefeld als Kombinationsfeld ein

    Dim objDB As DAO.Database
    Dim objTbl As DAO.TableDef
    Dim objFld As DAO.Field
    Dim objPrp As DAO.Property
    Dim strSQL As String

    Set objDB = CurrentDb
    Set objTbl = objDB.TableDefs(strChildTable)
    Set objFld = objTbl.Fields(strChildField)

    ' RowSource SQL generieren
    strSQL = "SELECT [" & strLookupKey & "], [" & strLookupDisplay & "] " & _
             "FROM [" & strLookupTable & "] ORDER BY [" & strLookupDisplay & "];"

    ' Alte Lookup-Properties entfernen
    On Error Resume Next
    objFld.Properties.Delete "RowSourceType"
    objFld.Properties.Delete "RowSource"
    objFld.Properties.Delete "BoundColumn"
    objFld.Properties.Delete "ColumnCount"
    objFld.Properties.Delete "ColumnWidths"
    objFld.Properties.Delete "ListWidth"
    objFld.Properties.Delete "DisplayControl"
    On Error GoTo 0

    ' Neue Lookup-Properties anlegen
    objFld.Properties.Append objFld.CreateProperty("RowSourceType", dbText, "Table/Query")
    objFld.Properties.Append objFld.CreateProperty("RowSource", dbMemo, strSQL)
    objFld.Properties.Append objFld.CreateProperty("BoundColumn", dbInteger, 1)
    objFld.Properties.Append objFld.CreateProperty("ColumnCount", dbInteger, 2)
    objFld.Properties.Append objFld.CreateProperty("ColumnWidths", dbText, "0cm;2cm")
    objFld.Properties.Append objFld.CreateProperty("ListWidth", dbText, "2cm")
    objFld.Properties.Append objFld.CreateProperty("DisplayControl", dbInteger, acComboBox)

    ' Objekte freigeben
    Set objFld = Nothing
    Set objTbl = Nothing
    Set objDB = Nothing
End Sub