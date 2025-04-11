Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub ExportToPlantUML()

    ' Exportiert alle .cls- und .bas-Komponenten in eine PlantUML-Datei
    ' Erstellt eine Klasse je Modul mit öffentlichen Properties und Methoden

    Dim strTargetFile As String
    Dim intFile As Integer
    Dim objComponent As Object
    Dim strLine As String
    Dim strCode As String
    Dim strComponentName As String
    Dim strComponentType As String
    Dim strMethod As String
    Dim strProperty As String

    strTargetFile = "C:\Export\AccessProject_ClassesAndModules.puml"
    intFile = FreeFile

    Open strTargetFile For Output As #intFile

    Print #intFile, "@startuml"
    Print #intFile, "skinparam classAttributeIconSize 0"
    Print #intFile, ""

    ' Durch alle Komponenten iterieren
    For Each objComponent In Application.VBE.VBProjects(1).VBComponents

        strComponentName = objComponent.Name
        strComponentType = vbNullString

        Select Case objComponent.Type
            Case 1 ' Standardmodul (.bas)
                strComponentType = "Module"
            Case 2 ' Klassenmodul (.cls)
                strComponentType = "Class"
            Case Else
                ' Andere Typen ignorieren
                GoTo NextComponent
        End Select

        Print #intFile, "class " & strComponentName & " {"

        ' Modultext durchgehen
        strCode = objComponent.CodeModule.Lines(1, objComponent.CodeModule.CountOfLines)
        Dim arrLines() As String
        arrLines = Split(strCode, vbCrLf)

        For Each strLine In arrLines

            ' Public Properties
            If strLine Like "Public * As *" Then
                Dim strPropParts() As String
                strPropParts = Split(Replace(strLine, "  ", " "), " ")
                If UBound(strPropParts) >= 3 Then
                    strProperty = "+" & strPropParts(3) & " " & strPropParts(1)
                    Print #intFile, "  " & strProperty
                End If
            End If

            ' Public Methods
            If strLine Like "Public Sub *" Or strLine Like "Public Function *" Then
                Dim strFuncParts() As String
                strFuncParts = Split(Replace(strLine, "  ", " "), " ")
                If UBound(strFuncParts) >= 2 Then
                    strMethod = "+" & strFuncParts(2) & "() : "
                    If InStr(1, strLine, "Function", vbTextCompare) > 0 Then
                        strMethod = strMethod & "any"
                    Else
                        strMethod = strMethod & "void"
                    End If
                    Print #intFile, "  " & strMethod
                End If
            End If

        Next strLine

        Print #intFile, "}" & vbCrLf

NextComponent:
    Next objComponent

    Print #intFile, "@enduml"
    Close #intFile

    MsgBox "PlantUML-Datei wurde exportiert:" & vbCrLf & strTargetFile, vbInformation

End Sub