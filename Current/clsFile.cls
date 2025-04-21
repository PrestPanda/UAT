Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Log As New clsLog


Public Enum enuFile_Type
    
    ftcls
    fttxt

End Enum

Public Sub Create_File(strFileName As String, strPath As String, strContent As String, FileType As enuFile_Type)

    Dim strFullPath As String
    Dim intFile As Integer
    
    ' Erstelle den vollständigen Dateipfad
    strFullPath = strPath & strFileName & Get_FileExtension_AsString_ByEnum(FileType)
    
    ' Datei öffnen (Nummer automatisch zuweisen)
    intFile = FreeFile
    Open strFullPath For Output As #intFile
    Print #intFile, strContent
    Close #intFile
    
    Log.WriteLine "Datei erstellt: " & strFullPath

End Sub
Public Function Get_FileExtension_AsString_ByEnum(FileType As enuFile_Type)

    Select Case FileType
    
        Case ftcls
            Get_FileExtension_AsString_ByEnum = ".cls"
        
        Case fttxt
            Get_FileExtension_AsString_ByEnum = ".txt"
            
    End Select
    
End Function