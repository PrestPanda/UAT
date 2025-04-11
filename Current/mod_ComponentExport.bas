Option Compare Database
Option Explicit

Dim Log As New clsLog

Public Sub ExportAllComponents()

    ' Exportiert alle Module, Klassen, Formulare, Berichte und Abfragen in zwei Verzeichnisse:
    ' 1. Archivordner mit Zeitstempel (inkl. Unterordner)
    ' 2. Aktueller flacher "Current"-Ordner ohne Unterverzeichnisse

    Dim strBasePath As String
    Dim strTimeStamp As String
    Dim strExportPath As String
    Dim strCurrentPath As String
    Dim strExportFile As String
    Dim strCurrentFile As String
    Dim objComponent As Object
    Dim objReport As AccessObject
    Dim objQuery As AccessObject
    Dim strName As String
    
    Log.WriteLine "Der Komponentenexport wurde gestartet."

    strBasePath = "C:\Users\Prest\Desktop\JCs Ultimate Access Tool\Versions\"  ' Basisverzeichnis
    strTimeStamp = Format(Now, "yyyy-mm-dd_hh-nn")
    strExportPath = strBasePath & "Export_" & strTimeStamp & "\"
    strCurrentPath = strBasePath & "Current\"

    ' Archivstruktur anlegen (für chronologische Versionen)
    CreateFolder strCurrentPath
    CreateFolder strExportPath
    CreateFolder strExportPath & "Modules\"
    CreateFolder strExportPath & "Classes\"
    CreateFolder strExportPath & "Forms\"
    CreateFolder strExportPath & "Reports\"
    CreateFolder strExportPath & "Queries\"

    ' "Current"-Ordner neu erstellen (löschen + neu)
    If Dir(strCurrentPath, vbDirectory) <> "" Then
        Kill strCurrentPath & "*.*"
    Else
        MkDir strCurrentPath
    End If

    ' VBA-Komponenten exportieren
    For Each objComponent In Application.VBE.VBProjects(1).VBComponents
        strName = objComponent.Name

        Select Case objComponent.Type
            Case 1 ' Modul
                strExportFile = strExportPath & "Modules\" & strName & ".bas"
                strCurrentFile = strCurrentPath & strName & ".bas"
                Application.SaveAsText acModule, strName, strExportFile
                Application.SaveAsText acModule, strName, strCurrentFile

            Case 2 ' Klassenmodul
                strExportFile = strExportPath & "Classes\" & strName & ".cls"
                strCurrentFile = strCurrentPath & strName & ".cls"
                Application.SaveAsText acModule, strName, strExportFile
                Application.SaveAsText acModule, strName, strCurrentFile

            Case 3 ' Formular
                strExportFile = strExportPath & "Forms\" & strName & ".frm"
                strCurrentFile = strCurrentPath & strName & ".frm"
                Application.SaveAsText acForm, strName, strExportFile
                Application.SaveAsText acForm, strName, strCurrentFile
        End Select
    Next objComponent

    ' Berichte exportieren
    For Each objReport In CurrentProject.AllReports
        strName = objReport.Name
        strExportFile = strExportPath & "Reports\" & strName & ".txt"
        strCurrentFile = strCurrentPath & strName & ".txt"
        Application.SaveAsText acReport, strName, strExportFile
        Application.SaveAsText acReport, strName, strCurrentFile
    Next objReport

    ' Abfragen exportieren (SQL)
    For Each objQuery In CurrentData.AllQueries
        strName = objQuery.Name
        strExportFile = strExportPath & "Queries\" & strName & ".sql"
        strCurrentFile = strCurrentPath & strName & ".sql"
        ExportQuerySQL strName, strExportFile
        ExportQuerySQL strName, strCurrentFile
    Next objQuery

    Log.WriteLine vbCrLf & vbCrLf & "Export abgeschlossen: " & vbCrLf & _
                    "Export_" & strTimeStamp & vbCrLf & _
                    "abgelegt unter: " & vbCrLf & _
                    strBasePath
    Log.WriteLine "'Current'-Ordner aktualisiert."
    Log.WriteEmptyLine


End Sub

Private Sub ExportQuerySQL(strQueryName As String, strFilePath As String)

    ' Exportiert die SQL-Definition einer Abfrage in eine .sql-Datei

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim intFile As Integer

    Set db = CurrentDb
    Set qdf = db.QueryDefs(strQueryName)

    intFile = FreeFile
    Open strFilePath For Output As intFile
    Print #intFile, qdf.SQL
    Close intFile

End Sub
Public Sub CreateFolder(strFolderPath As String)

    ' Erstellt einen Ordner, wenn dieser noch nicht existiert

    If Dir(strFolderPath, vbDirectory) = "" Then
        MkDir strFolderPath
    End If

End Sub