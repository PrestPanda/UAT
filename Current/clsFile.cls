Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Sub Create_File_txt(strFileName As String, strPath As String, strContent As String)
    ' Erstellt eine Textdatei mit dem angegebenen Namen, Pfad und Inhalt
    
    Dim strFullPath As String
    Dim intFile As Integer
    
    ' Erstelle den vollständigen Dateipfad
    strFullPath = strPath & strFileName & ".txt"
    
    ' Datei öffnen (Nummer automatisch zuweisen)
    intFile = FreeFile
    Open strFullPath For Output As #intFile
    
    ' Inhalt in die Datei schreiben
    Print #intFile, strContent
    
    ' Datei schließen
    Close #intFile
    
    ' Meldung für erfolgreiche Erstellung (optional)
    Debug.Print "Datei erstellt: " & strFullPath
    
End Sub
Public Sub Create_File_cls(strFileName As String, strPath As String, strContent As String)
    ' Erstellt eine Textdatei mit dem angegebenen Namen, Pfad und Inhalt
    
    Dim strFullPath As String
    Dim intFile As Integer
    
    ' Erstelle den vollständigen Dateipfad
    strFullPath = strPath & strFileName & ".cls"
    
    ' Datei öffnen (Nummer automatisch zuweisen)
    intFile = FreeFile
    Open strFullPath For Output As #intFile
    
    ' Inhalt in die Datei schreiben
    Print #intFile, strContent
    
    ' Datei schließen
    Close #intFile
    
    ' Meldung für erfolgreiche Erstellung (optional)
    Debug.Print "Datei erstellt: " & strFullPath
    
End Sub