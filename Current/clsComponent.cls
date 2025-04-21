Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum enuComponent_Types

    ctTable
    ctQuery
    ctForm
    ctModule
    ctClassModule
    
End Enum


Dim Log As New clsLog

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

Public Sub ExportAllComponents()

    ' Exportiert alle Module, Klassen, Formulare, Berichte und Abfragen in zwei Verzeichnisse:
    ' 1. Archivordner mit Zeitstempel (inkl. Unterordner)
    ' 2. Aktueller flacher "Current"-Ordner ohne Unterverzeichnisse
    
    
    Log.WriteLine "Der Komponentenexport wurde gestartet."

    strBasePath = "C:\Users\Prest\Desktop\JCs Ultimate Access Tool\Versions\"  ' Basisverzeichnis
    strTimeStamp = Format(Now, "yyyy-mm-dd_hh-nn")
    strExportPath = strBasePath & "Export_" & strTimeStamp & "\"
    strCurrentPath = strBasePath & "Current\"

    Export_Component_Directory_Create
    Export_Component_Directory_Current_Clear
    

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

            Case 100 ' Formular-Modul
                Export_Component_Form strName, strCurrentPath
                Export_Component_Form strName, strExportPath & "Forms\"
                
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
                    strBasePath & vbCrLf
    Log.WriteLine "'Current'-Ordner aktualisiert."
    Log.WriteEmptyLine


End Sub
Private Sub Export_Component_Directory_Create()

    Dim Directory As New clsDirectory

 ' Archivstruktur anlegen (für chronologische Versionen)
    Directory.CreateFolder strCurrentPath
    Directory.CreateFolder strExportPath
    Directory.CreateFolder strExportPath & "Modules\"
    Directory.CreateFolder strExportPath & "Classes\"
    Directory.CreateFolder strExportPath & "Forms\"
    Directory.CreateFolder strExportPath & "Reports\"
    Directory.CreateFolder strExportPath & "Queries\"


End Sub
Private Sub Export_Component_Directory_Current_Clear()

    ' "Current"-Ordner neu erstellen (löschen + neu)
    If Dir(strCurrentPath, vbDirectory) <> "" Then
        Kill strCurrentPath & "*.*"
    Else
        MkDir strCurrentPath
    End If

End Sub
Public Sub Export_Component_Form(strFormName As String, strExportPath As String)

    ' Exportiert alle Steuerelemente eines Formulars in eine TXT-Datei
    ' sowie den VBA-Code des Formulars als BAS-Datei

    
    Dim objForm As Access.Form
    Dim objControl As Control
    Dim intFile As Integer
    Dim strFileTxt As String
    Dim strFileBas As String
    Dim objComponent As Object

    strFormName = Replace(strFormName, "Form_", "")

    ' Pfade definieren
    strFileTxt = strExportPath & "\" & strFormName & "_Controls.txt"
    strFileBas = strExportPath & "\" & strFormName & ".bas"

    ' Formular im Entwurf öffnen (falls nicht bereits)
    DoCmd.OpenForm strFormName, acDesign, WindowMode:=acHidden
    Set objForm = Forms(strFormName)

    ' Alle Steuerelemente exportieren
    intFile = FreeFile
    Open strFileTxt For Output As #intFile

    Print #intFile, "Form: " & strFormName
    Print #intFile, "------------------------------------"
    For Each objControl In objForm.Controls
        Print #intFile, "Name: " & objControl.Name & vbTab & "Typ: " & TypeName(objControl)
    Next objControl

    Close #intFile

    DoCmd.Close acForm, strFormName, acSaveNo

    ' VBA-Code exportieren
    Set objComponent = Application.VBE.VBProjects(1).VBComponents("Form_" & strFormName)
    objComponent.Export strFileBas

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