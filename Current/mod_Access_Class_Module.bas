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

    ' Gibt ein Array mit eindeutigen Property-Namen zurück (ohne Datentyp).
    ' Jede Property wird nur einmal aufgenommen (egal ob Get/Let/Set).

    Dim objModule As VBIDE.CodeModule
    Dim lngLine As Long
    Dim lngLastLine As Long
    Dim strLineText As String
    Dim colProperties As Collection
    Dim arrResult() As String
    Dim i As Long

    Set colProperties = New Collection

    Set objModule = Application.VBE.ActiveVBProject.VBComponents(strClassName).CodeModule
    lngLastLine = objModule.CountOfLines

    For lngLine = 1 To lngLastLine
        strLineText = Trim(objModule.Lines(lngLine, 1))
        
        ' Prüft auf Property Get / Let / Set
        If (InStr(1, strLineText, "Property Get", vbTextCompare) > 0) Or _
           (InStr(1, strLineText, "Property Let", vbTextCompare) > 0) Or _
           (InStr(1, strLineText, "Property Set", vbTextCompare) > 0) Then
           
            Dim strPropertyName As String
            Dim lngPosNameStart As Long
            Dim lngPosAs As Long
            
            ' Property-Name extrahieren
            lngPosNameStart = InStr(1, strLineText, "Property", vbTextCompare)
            lngPosNameStart = InStr(lngPosNameStart + 8, strLineText, " ", vbTextCompare)
            lngPosAs = InStr(1, strLineText, "As", vbTextCompare)
            
            If lngPosNameStart > 0 And lngPosAs > 0 Then
                strPropertyName = Trim(Mid(strLineText, lngPosNameStart + 1, lngPosAs - lngPosNameStart - 1))
                
                ' Prüfen, ob Property schon in der Collection enthalten ist (einzigartig)
                Dim blnExists As Boolean
                blnExists = False
                For i = 1 To colProperties.Count
                    If colProperties(i) = strPropertyName Then
                        blnExists = True
                        Exit For
                    End If
                Next i

                ' Wenn noch nicht enthalten, hinzufügen
                If Not blnExists Then
                    colProperties.Add strPropertyName
                End If
            End If

        End If
    Next lngLine

    ' Ergebnis-Array vorbereiten
    If colProperties.Count > 0 Then
        ReDim arrResult(0 To colProperties.Count - 1)
        For i = 1 To colProperties.Count
            arrResult(i - 1) = colProperties(i)
        Next i
        Access_Class_Module_Get_PropertyNames = arrResult
    Else
        Access_Class_Module_Get_PropertyNames = Null
    End If

End Function

'Public Function Access_Class_Module_Get_PropertyNamesAndTypes(strClassName As String) As Variant
'
'    ' Gibt ein Array mit Property-Namen und deren Datentyp zurück.
'    ' Jedes Element ist ein 2D-Array: (0) = Name, (1) = Datentyp
'
'    Dim objModule As VBIDE.CodeModule
'    Dim lngLine As Long
'    Dim lngLastLine As Long
'    Dim strLineText As String
'    Dim colProperties As Collection
'    Dim arrResult() As String
'    Dim i As Long
'
'    Set colProperties = New Collection
'
'    Set objModule = Application.VBE.ActiveVBProject.VBComponents(strClassName).CodeModule
'    lngLastLine = objModule.CountOfLines
'
'    For lngLine = 1 To lngLastLine
'        strLineText = Trim(objModule.Lines(lngLine, 1))
'
'        ' Prüft auf Property Get oder Let/Set
'        If (InStr(1, strLineText, "Property Get", vbTextCompare) > 0) Or _
'           (InStr(1, strLineText, "Property Let", vbTextCompare) > 0) Or _
'           (InStr(1, strLineText, "Property Set", vbTextCompare) > 0) Then
'
'            Dim strPropertyName As String
'            Dim strDataType As String
'            Dim lngPosNameStart As Long
'            Dim lngPosAs As Long
'
'            ' Property-Name extrahieren
'            lngPosNameStart = InStr(1, strLineText, "Property", vbTextCompare)
'            lngPosNameStart = InStr(lngPosNameStart + 8, strLineText, " ", vbTextCompare)
'            lngPosAs = InStr(1, strLineText, "As", vbTextCompare)
'
'            If lngPosNameStart > 0 And lngPosAs > 0 Then
'                strPropertyName = Trim(Mid(strLineText, lngPosNameStart + 1, lngPosAs - lngPosNameStart - 1))
'                strDataType = Trim(Mid(strLineText, lngPosAs + 2))
'
'                ' Zeilenumbruch entfernen, falls vorhanden
'                strDataType = Replace(strDataType, vbCr, "")
'                strDataType = Replace(strDataType, vbLf, "")
'
'                ' Füge als Array (Name, Datentyp) zur Collection hinzu
'                Dim arrProperty(1) As String
'                arrProperty(0) = strPropertyName
'                arrProperty(1) = strDataType
'                colProperties.Add arrProperty
'            End If
'
'        End If
'    Next lngLine
'
'    ' Ergebnis-Array vorbereiten
'    If colProperties.Count > 0 Then
'        ReDim arrResult(0 To colProperties.Count - 1, 0 To 1)
'        For i = 1 To colProperties.Count
'            arrResult(i - 1, 0) = colProperties(i)(0) ' Name
'            arrResult(i - 1, 1) = colProperties(i)(1) ' Datentyp
'        Next i
'        Access_Class_Module_Get_PropertyNamesAndTypes = arrResult
'    Else
'        Access_Class_Module_Get_PropertyNamesAndTypes = Null
'    End If
'
'End Function

'Public Function Access_Class_Module_Get_PropertyNames(strClassName As String) As Variant
'
'     'Gibt ein Array mit allen Property-Namen aus dem angegebenen Klassenmodul zurück
'
'    Dim objComponent As VBIDE.VBComponent
'    Dim objCodeModule As VBIDE.CodeModule
'    Dim lngLineCount As Long
'    Dim lngLine As Long
'    Dim strLineText As String
'    Dim colProperties As Collection
'    Dim strPropertyName As String
'    Dim varResult() As String
'    Dim i As Long
'
'    Set objComponent = Application.VBE.ActiveVBProject.VBComponents(strClassName)
'    Set objCodeModule = objComponent.CodeModule
'    Set colProperties = New Collection
'
'    lngLineCount = objCodeModule.CountOfLines
'
'    For lngLine = 1 To lngLineCount
'        strLineText = Trim(objCodeModule.Lines(lngLine, 1))
'
'        If strLineText Like "Public Property *" Or strLineText Like "Private Property *" Then
'            strPropertyName = GetPropertyNameFromLine(strLineText)
'            If Len(strPropertyName) > 0 Then
'                On Error Resume Next
'                colProperties.Add strPropertyName, strPropertyName ' doppelte ignorieren
'                On Error GoTo 0
'            End If
'        End If
'    Next lngLine
'
'    If colProperties.Count > 0 Then
'        ReDim varResult(0 To colProperties.Count - 1)
'        For i = 1 To colProperties.Count
'            varResult(i - 1) = colProperties(i)
'        Next i
'        Access_Class_Module_Get_PropertyNames = varResult
'    Else
'        Access_Class_Module_Get_PropertyNames = Null
'    End If
'
'
'End Function
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