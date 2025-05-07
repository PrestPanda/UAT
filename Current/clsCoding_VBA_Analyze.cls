Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Function Get_CodeElements_Details_AsArray(strModuleName As String) As Variant
    ' Gibt für jedes Code-Element Name, Parameter, Rückgabetyp und Typ zurück

    Dim strCode As String
    Dim varElements As Variant
    Dim lCount As Long
    Dim varResult As Variant
    Dim lIndex As Long
    Dim varParams As Variant

    ' Gesamten Code des Moduls ermitteln
    strCode = Get_CodeModule_AsString(strModuleName)

    ' Code-Elemente als Array holen
    varElements = Get_CodeElements_AsArray(strCode)
    lCount = UBound(varElements) - LBound(varElements) + 1

    ' Ergebnis-Array: Zeilen = Elemente, Spalten = 0:Name, 1:Parameter, 2:ReturnType, 3:Type
    ReDim varResult(0 To lCount - 1, 0 To 3)

    For lIndex = LBound(varElements) To UBound(varElements)
        varParams = Get_CodeElement_Parameters_AsArray(varElements(lIndex))
        varResult(lIndex, 0) = GetCodeElement_Name_AsString(varElements(lIndex))
        varResult(lIndex, 1) = varParams
        varResult(lIndex, 2) = Get_CodeElement_ReturnType_AsString(varElements(lIndex))
        varResult(lIndex, 3) = Get_CodeElement_Type(varElements(lIndex))
    Next lIndex

    Get_CodeElements_Details_AsArray = varResult


End Function
Public Function Get_CodeModule_AsString(strModuleName As String) As String
    ' Gibt den gesamten Code eines Klassenmoduls zurück

    Dim objVBComp As VBIDE.VBComponent
    Dim strCode As String

    Set objVBComp = Application.VBE.ActiveVBProject.VBComponents(strModuleName)
    strCode = objVBComp.CodeModule.Lines(1, objVBComp.CodeModule.CountOfLines)

    Get_CodeModule_AsString = strCode

End Function

Public Function Get_CodeElements_AsArray(ByVal strModuleCode As String) As Variant
    ' Gibt ein 1D Array zurück, welches den aufgeteilten Code enthält, inkl. Visibility-Präfix

    Dim varLines As Variant
    Dim lngStarts() As Long
    Dim varElements As Variant
    Dim strTrimLine As String
    Dim strLow As String
    Dim arrQualifiers As Variant
    Dim vQualifier As Variant
    Dim strCodeBlock As String
    Dim lngCount As Long
    Dim lngLineIndex As Long
    Dim lngIdx As Long

    arrQualifiers = Array("public ", "private ", "friend ", "static ")
    varLines = Split(strModuleCode, vbCrLf)

    ' Start-Positionen der Code-Elemente ermitteln
    For lngLineIndex = LBound(varLines) To UBound(varLines)
        strTrimLine = Trim(varLines(lngLineIndex))
        strLow = LCase$(strTrimLine)
        For Each vQualifier In arrQualifiers
            Do While Left$(strLow, Len(vQualifier)) = vQualifier
                strLow = Mid$(strLow, Len(vQualifier) + 1)
            Loop
        Next vQualifier
        If Left$(strLow, 4) = "sub " Or Left$(strLow, 9) = "function " Or Left$(strLow, 9) = "property " Then
            If lngCount = 0 Then
                ReDim lngStarts(0)
            Else
                ReDim Preserve lngStarts(0 To lngCount)
            End If
            lngStarts(lngCount) = lngLineIndex
            lngCount = lngCount + 1
        End If
    Next lngLineIndex

    ' Keine Elemente gefunden: leeres Array zurückgeben
    If lngCount = 0 Then
        Get_CodeElements_AsArray = Array()
        Exit Function
    End If

    ' Array für Codeblöcke dimensionieren
    ReDim varElements(0 To lngCount - 1)

    ' Codeblöcke zusammenstellen
    For lngIdx = 0 To lngCount - 1
        strCodeBlock = vbNullString
        Dim lngEndLine As Long
        If lngIdx < lngCount - 1 Then
            lngEndLine = lngStarts(lngIdx + 1) - 1
        Else
            lngEndLine = UBound(varLines)
        End If
        For lngLineIndex = lngStarts(lngIdx) To lngEndLine
            strCodeBlock = strCodeBlock & varLines(lngLineIndex) & vbCrLf
        Next lngLineIndex
        varElements(lngIdx) = strCodeBlock
    Next lngIdx

    Get_CodeElements_AsArray = varElements
End Function

Public Function GetCodeElement_Name_AsString(ByVal strElementCode As String) As String
    ' Gibt den Namen des Elements als String zurück, inkl. Property Get/Let/Set

    Dim varLines As Variant
    Dim strSignature As String
    Dim sigLow As String
    Dim lPropPos As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim strAfter As String
    Dim arrParts As Variant
    Dim strName As String

    varLines = Split(strElementCode, vbCrLf)
    If UBound(varLines) < 0 Then
        GetCodeElement_Name_AsString = ""
        Exit Function
    End If

    strSignature = Trim$(varLines(0))
    sigLow = LCase$(strSignature)

    ' Position der Keywords per InStr ermitteln
    If InStr(1, sigLow, "sub ", vbTextCompare) > 0 Then
        lStart = InStr(1, sigLow, "sub ", vbTextCompare) + Len("sub ")
    ElseIf InStr(1, sigLow, "function ", vbTextCompare) > 0 Then
        lStart = InStr(1, sigLow, "function ", vbTextCompare) + Len("function ")
    ElseIf InStr(1, sigLow, "property ", vbTextCompare) > 0 Then
        lPropPos = InStr(1, sigLow, "property ", vbTextCompare) + Len("property ")
        strAfter = Mid$(strSignature, lPropPos)
        arrParts = Split(Trim$(strAfter), " ")
        If UBound(arrParts) >= 1 Then
            strName = arrParts(1)
        Else
            strName = arrParts(0)
        End If
        ' Klammern abschneiden
        lEnd = InStr(strName, "(")
        If lEnd > 0 Then strName = Left$(strName, lEnd - 1)

        GetCodeElement_Name_AsString = strName
        Exit Function
    Else
        GetCodeElement_Name_AsString = ""
        Exit Function
    End If

    ' Sub/Function Name extrahieren
    lEnd = lStart
    Do While lEnd <= Len(strSignature) And Mid$(strSignature, lEnd, 1) Like "[A-Za-z0-9_]"
        lEnd = lEnd + 1
    Loop

    strName = Mid$(strSignature, lStart, lEnd - lStart)
    GetCodeElement_Name_AsString = strName
End Function


Public Function Get_CodeElement_Parameters_AsArray(ByVal strElementCode As String) As Variant
    ' Gibt ein Array zurück, welches die Namen und Datentypen der übergebenen Parameter enthält

    Dim varLines As Variant
    Dim strSignature As String
    Dim lOpen As Long, lClose As Long
    Dim strParams As String
    Dim varParams As Variant
    Dim i As Long

    varLines = Split(strElementCode, vbCrLf)
    If UBound(varLines) < 0 Then
        Get_CodeElement_Parameters_AsArray = Array()
        Exit Function
    End If
    strSignature = Trim(varLines(0))

    lOpen = InStr(1, strSignature, "(", vbTextCompare)
    lClose = InStrRev(strSignature, ")")
    If lOpen > 0 And lClose > lOpen Then
        strParams = Mid$(strSignature, lOpen + 1, lClose - lOpen - 1)
        If Trim$(strParams) = "" Then
            Get_CodeElement_Parameters_AsArray = Array()
            Exit Function
        End If
        varParams = Split(strParams, ",")
        For i = LBound(varParams) To UBound(varParams)
            varParams(i) = Trim$(varParams(i))
        Next i
        Get_CodeElement_Parameters_AsArray = varParams
    Else
        Get_CodeElement_Parameters_AsArray = Array()
    End If
End Function

Public Function Get_CodeElement_ReturnType_AsString(ByVal strElementCode As String) As String
    ' Gibt den Datentyp des Rückgabewerts zurück, sofern vorhanden

    Dim varLines As Variant
    Dim strSignature As String
    Dim lAs As Long
    Dim strReturnType As String
    Dim sigLow As String

    varLines = Split(strElementCode, vbCrLf)
    If UBound(varLines) < 0 Then
        Get_CodeElement_ReturnType_AsString = ""
        Exit Function
    End If
    strSignature = Trim(varLines(0))
    sigLow = LCase$(strSignature)

    If Left$(sigLow, 12) = "property let " Then
        ' Parameter-Datentyp für Property Let
        Dim lOpen As Long, lClose As Long, strParams As String
        lOpen = InStr(1, strSignature, "(")
        lClose = InStr(strSignature, ")")
        If lOpen > 0 And lClose > lOpen Then
            strParams = Mid$(strSignature, lOpen + 1, lClose - lOpen - 1)
            lAs = InStrRev(strParams, " As ", , vbTextCompare)
            If lAs > 0 Then
                strReturnType = Trim$(Mid$(strParams, lAs + Len(" As ")))
            End If
        End If
    Else
        ' Standard-Rückgabetyp
        lAs = InStrRev(strSignature, " As ", , vbTextCompare)
        If lAs > 0 Then
            strReturnType = Trim$(Mid$(strSignature, lAs + Len(" As ")))
            strReturnType = Replace(strReturnType, ")", "")
        End If
    End If

    Get_CodeElement_ReturnType_AsString = strReturnType
End Function

Public Function Get_CodeElement_Type(ByVal strElementCode As String) As String
    ' Gibt Sichtbarkeit und Typ eines Elements zurück, z.B. "Public Property"

    Dim varLines As Variant
    Dim strSignature As String
    Dim strLow As String
    Dim arrQualifiers As Variant
    Dim vQualifier As Variant
    Dim strVisibility As String
    Dim strType As String

    arrQualifiers = Array("Public ", "Private ", "Friend ", "Static ")
    varLines = Split(strElementCode, vbCrLf)
    If UBound(varLines) < 0 Then
        Get_CodeElement_Type = ""
        Exit Function
    End If

    strSignature = Trim$(varLines(0))
    strLow = LCase$(strSignature)

    ' Sichtbarkeitsqualifier bestimmen
    strVisibility = ""
    For Each vQualifier In arrQualifiers
        If LCase$(Left$(strSignature, Len(vQualifier))) = LCase$(vQualifier) Then
            strVisibility = Trim$(vQualifier)
            Exit For
        End If
    Next vQualifier

    ' Typ per InStr ermitteln, unabhängig von Position
    If InStr(1, strLow, "property ", vbTextCompare) > 0 Then
        strType = "Property"
    ElseIf InStr(1, strLow, "function ", vbTextCompare) > 0 Then
        strType = "Function"
    ElseIf InStr(1, strLow, "sub ", vbTextCompare) > 0 Then
        strType = "Sub"
    Else
        strType = ""
    End If



    Get_CodeElement_Type = strType

End Function