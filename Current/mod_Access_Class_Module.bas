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

Public Function Access_Class_Module_Get_PropertyNames(strClassName As String) As Variant

    ' Gibt ein Array mit allen Property-Namen aus dem angegebenen Klassenmodul zurück

    Dim objComponent As VBIDE.VBComponent
    Dim objCodeModule As VBIDE.CodeModule
    Dim lngLineCount As Long
    Dim lngLine As Long
    Dim strLineText As String
    Dim colProperties As Collection
    Dim strPropertyName As String
    Dim varResult() As String
    Dim i As Long

    Set objComponent = Application.VBE.ActiveVBProject.VBComponents(strClassName)
    Set objCodeModule = objComponent.CodeModule
    Set colProperties = New Collection

    lngLineCount = objCodeModule.CountOfLines

    For lngLine = 1 To lngLineCount
        strLineText = Trim(objCodeModule.Lines(lngLine, 1))
        
        If strLineText Like "Public Property *" Or strLineText Like "Private Property *" Then
            strPropertyName = GetPropertyNameFromLine(strLineText)
            If Len(strPropertyName) > 0 Then
                On Error Resume Next
                colProperties.Add strPropertyName, strPropertyName ' doppelte ignorieren
                On Error GoTo 0
            End If
        End If
    Next lngLine

    If colProperties.Count > 0 Then
        ReDim varResult(0 To colProperties.Count - 1)
        For i = 1 To colProperties.Count
            varResult(i - 1) = colProperties(i)
        Next i
        Access_Class_Module_Get_PropertyNames = varResult
    Else
        Access_Class_Module_Get_PropertyNames = Null
    End If

    
End Function
Private Function GetPropertyNameFromLine(strLine As String) As String
'To-Do: Rename

    ' Extrahiert den Property-Namen aus einer Zeile wie "Property Get/Let/Set Name(...)"

    Dim strParts() As String
    strParts = Split(Trim(strLine), " ")
    
    If UBound(strParts) >= 2 Then
        GetPropertyNameFromLine = Trim$(Split(strParts(3), "(")(0))
    Else
        GetPropertyNameFromLine = ""
    End If

    
End Function
Public Function Access_Class_Module_Get_MethodNames(strClassName As String) As Variant

    ' Gibt ein Array mit allen Methoden (Sub/Function) eines Klassenmoduls zurück – unabhängig von Sichtbarkeit

    Dim objComponent As VBIDE.VBComponent
    Dim objCodeModule As VBIDE.CodeModule
    Dim lngLineCount As Long
    Dim lngLine As Long
    Dim strLineText As String
    Dim colMethods As Collection
    Dim varResult() As String
    Dim i As Long

    Set objComponent = Application.VBE.ActiveVBProject.VBComponents(strClassName)
    Set objCodeModule = objComponent.CodeModule
    Set colMethods = New Collection

    lngLineCount = objCodeModule.CountOfLines

    For lngLine = 1 To lngLineCount
        strLineText = Trim(objCodeModule.Lines(lngLine, 1))

        If strLineText Like "*Sub *" Or strLineText Like "*Function *" Then
            Dim strMethodName As String
            strMethodName = GetMethodNameFromLine(strLineText)
            If Len(strMethodName) > 0 Then
                On Error Resume Next
                colMethods.Add strMethodName, strMethodName
                On Error GoTo 0
            End If
        End If
    Next lngLine

    If colMethods.Count > 0 Then
        ReDim varResult(0 To colMethods.Count - 1)
        For i = 1 To colMethods.Count
            varResult(i - 1) = colMethods(i)
        Next i
        Access_Class_Module_Get_MethodNames = varResult
    Else
        Access_Class_Module_Get_MethodNames = Null
    End If

    
End Function
Private Function GetMethodNameFromLine(strLine As String) As String

    ' Extrahiert den Methodennamen aus einer Zeile wie "Public Sub Foo(...)" oder "Private Function Bar(...)"

    Dim strParts() As String
    strParts = Split(Trim(strLine), " ")

    Dim i As Long
    For i = LBound(strParts) To UBound(strParts) - 1
        If LCase(strParts(i)) = "sub" Or LCase(strParts(i)) = "function" Then
            GetMethodNameFromLine = Split(strParts(i + 1), "(")(0)
            Exit Function
        End If
    Next i

    GetMethodNameFromLine = ""
    
End Function