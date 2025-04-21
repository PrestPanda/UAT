Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub CreateFolder(strFolderPath As String)

    ' Erstellt einen Ordner, wenn dieser noch nicht existiert

    If Dir(strFolderPath, vbDirectory) = "" Then
        MkDir strFolderPath
    End If

End Sub