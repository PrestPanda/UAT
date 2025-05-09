Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Math As New clsmath

Public Enum enuWriting_Object
    pboLog
    pboDebug
    pboVBE
End Enum
Public Enum enuWriting_ProgressBar_Design
    pbdesHash
    pbdesPointer
    pbdesDot
End Enum
Public Enum enu_Writing_Separator_Design
    sepdesDash
    sepdesEqual
    sepdesStars
    sepdesCustom
End Enum
Public Enum enuWriting_Prefix_Type
    ptInfo
    ptError
    ptWarning
End Enum

'############################################ CONFIGURATION #######################################################
Private Function Get_ProgressBar_StandardLength(CallingObject As enuWriting_Object) As Long
'Gibt die Standardl�nge f�r bestimmte Objekte an z.B. clsLog, Debug-Fenster

    Select Case CallingObject
    
        Case pboDebug
            Get_ProgressBar_StandardLength = 40
        Case Else
            Get_ProgressBar_StandardLength = 0
            
    End Select

End Function
Private Function Get_Separator_StandardLength(CallingObject As enuWriting_Object) As Long
'Gibt die Standardl�nge f�r bestimmte Objekte an z.B. clsLog, Debug-Fenster

    Select Case CallingObject
    
        Case pboDebug
            Get_Separator_StandardLength = 40
        Case Else
            Get_Separator_StandardLength = 0
            
    End Select

End Function


'############################################ PROGRESSBAR #######################################################
Public Function Get_ProgressBar(lngItemsCount As Long, lngItemsCount_Total As Long, _
                                CallingObject As enuWriting_Object, _
                                Design As enuWriting_ProgressBar_Design) As String
     ' Erstellt eine textbasierte Fortschrittsanzeige basierend auf der Anzahl der bearbeiteten Elemente

    Dim lngPercentage As Long
    Dim lngFilled As Long
    Dim strBar As String
    Dim lngTotalLength As Long
    
    lngTotalLength = Get_ProgressBar_StandardLength(CallingObject) ' L�nge der Fortschrittsleiste
    lngPercentage = Math.CalculatePercentage(lngItemsCount, lngItemsCount_Total)
    
    ' Berechnet die Anzahl der gef�llten Bl�cke
    lngFilled = (lngPercentage * lngTotalLength) \ 100
        
    Select Case Design
        
        Case pbdesDot
            Get_ProgressBar = Get_ProgressBar_Dots(lngFilled, lngTotalLength, lngPercentage)
            
        Case pbdesHash
            Get_ProgressBar = Get_ProgressBar_Hash(lngFilled, lngTotalLength, lngPercentage)
            
        Case pbdesPointer
            Get_ProgressBar = Get_ProgressBar_Pointer(lngFilled, lngTotalLength, lngPercentage)
            
    End Select


End Function
Private Function Get_ProgressBar_Hash(lngFilled As Long, lngTotalLenght As Long, lngPercentage As Long) As String
' Erstellt die Fortschrittsleiste mit Hashes

    Get_ProgressBar_Hash = "[" & String(lngFilled, "#") & String(lngTotalLenght - lngFilled, "-") & "] " & lngPercentage & "%"
    
End Function

Private Function Get_ProgressBar_Pointer(lngFilled As Long, lngTotalLenght As Long, lngPercentage As Long) As String
' Erstellt die Fortschrittsleiste mit einem Pfeil

    Get_ProgressBar_Pointer = "[" & String(lngFilled, "=") & ">" & String(lngTotalLenght - lngFilled - 1, "-") & "] " & lngPercentage & "%"
    
End Function

Private Function Get_ProgressBar_Dots(lngFilled As Long, lngTotalLenght As Long, lngPercentage As Long) As String
' Erstellt die Fortschrittsleiste mit Punkten

    Get_ProgressBar_Dots = "[" & String(lngFilled, ".") & String(lngTotalLenght - lngFilled, "-") & "] " & lngPercentage & "%"
    
End Function




'############################################ SEPARATOR #######################################################
Public Function Get_Separator(CallingObject As enuWriting_Object, Design As enu_Writing_Separator_Design, _
                                Optional strCustom As String) As String
' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen in der angegebenen L�nge

    Dim lngTotalLength As Long
    
    lngTotalLength = Get_ProgressBar_StandardLength(CallingObject) ' L�nge der Fortschrittsleiste
        
    Select Case Design
        
        Case sepdesCustom
            If strCustom <> "" Then
                Get_Separator = Get_Separator_Custom(lngTotalLength, strCustom)
            Else
                Get_Separator = ""
            End If
            
        Case sepdesDash
            Get_Separator = Get_Separator_Dashes(lngTotalLength)
            
        Case sepdesEqual
            Get_Separator = Get_Separator_Equals(lngTotalLength)
            
         Case sepdesStars
            Get_Separator = Get_Separator_Stars(lngTotalLength)
            
    End Select


End Function


Private Function Get_Separator_Dashes(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Bindestrichen (-) in der angegebenen L�nge
    
    Get_Separator_Dashes = String(lngTotalLength, "-")
    
End Function

Private Function Get_Separator_Equals(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Gleichheitszeichen (=) in der angegebenen L�nge
    
    Get_Separator_Equals = String(lngTotalLength, "=")
    
End Function

Private Function Get_Separator_Stars(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Sternen (*) in der angegebenen L�nge
    
    Get_Separator_Stars = String(lngTotalLength, "*")
    
End Function

Private Function Get_Separator_Custom(lngTotalLength As Long, strChar As String) As String
    ' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen in der angegebenen L�nge
    
        Get_Separator_Custom = String((lngTotalLength / Len(strChar)), strChar)

End Function

Public Function Get_Separator_WithText(strText As String, CallingObject As enuWriting_Object, _
                                        Design As enu_Writing_Separator_Design, _
                                        Optional strChar As String, Optional blnUpperCase As Boolean = False) As String
' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen und einem mittigen Text
    
    Dim strFormattedText As String
    Dim lngTextLength As Long
    Dim lngSideLength As Long
    Dim lngTotalLength As Long
    
    lngTotalLength = Get_Separator_StandardLength(CallingObject)
    
    ' Text wahlweise in Gro�buchstaben umwandeln
    If blnUpperCase Then
        strFormattedText = UCase(strText)
    Else
        strFormattedText = strText
    End If
    
    ' Berechnung der Seitenl�ngen
    lngTextLength = Len(strFormattedText) + 2 ' +2 f�r die zus�tzlichen Leerzeichen
    lngSideLength = (lngTotalLength - lngTextLength) \ 2
    
    ' Falls die Gesamtl�nge nicht exakt aufgeteilt werden kann, eine Seite um 1 Zeichen erweitern
    If (lngTotalLength - lngTextLength) Mod 2 <> 0 Then
        lngSideLength = lngSideLength + 1
    End If
    
    Select Case Design
    
        Case sepdesDash
            Get_Separator_WithText = Get_Separator_WithText_Dashes(lngSideLength) & " " & strFormattedText & " " & Get_Separator_WithText_Dashes(lngSideLength)
        
        Case sepdesEqual
            Get_Separator_WithText = Get_Separator_WithText_Equals(lngSideLength) & " " & strFormattedText & " " & Get_Separator_WithText_Equals(lngSideLength)
        
        Case sepdesStars
            Get_Separator_WithText = Get_Separator_WithText_Stars(lngSideLength) & " " & strFormattedText & " " & Get_Separator_WithText_Stars(lngSideLength)
        
        Case sepdesCustom
            Get_Separator_WithText = Get_Separator_WithText_Custom(lngSideLength, strChar) & " " & strFormattedText & " " & Get_Separator_WithText_Custom(lngSideLength, strChar)
    
    End Select
    
End Function

Private Function Get_Separator_WithText_Dashes(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Bindestrichen (-) in der angegebenen L�nge
    Get_Separator_WithText_Dashes = String(lngTotalLength, "-")
End Function

Private Function Get_Separator_WithText_Equals(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Gleichheitszeichen (=) in der angegebenen L�nge
    Get_Separator_WithText_Equals = String(lngTotalLength, "=")
End Function

Private Function Get_Separator_WithText_Stars(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Sternen (*) in der angegebenen L�nge
    Get_Separator_WithText_Stars = String(lngTotalLength, "*")
End Function

Private Function Get_Separator_WithText_Custom(lngTotalLength As Long, strChar As String) As String
    ' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen in der angegebenen L�nge
    Get_Separator_WithText_Custom = String((lngTotalLength \ Len(strChar)), strChar)
End Function
'############################################ TIMESTAMP #######################################################
Public Function Get_Timestamp() As String
    ' Gibt den aktuellen Zeitstempel zur�ck
    Get_Timestamp = "[" & Format(Now, "yyyy-mm-dd HH:NN:SS") & "] "
End Function
'############################################ PREFIX #######################################################
Public Function Get_Prefix_EachLine_Apostrophe(strInput As String) As String
    ' F�gt jeder Zeile des �bergebenen Strings ein Apostroph voran.
    ' Gibt den modifizierten String zur�ck, der nahtlos in den VBE eingef�gt werden kann.

    Dim arrLines() As String
    Dim strResult As String
    Dim lngIndex As Long

    arrLines = Split(strInput, vbCrLf)

    For lngIndex = LBound(arrLines) To UBound(arrLines)
        arrLines(lngIndex) = "'" & arrLines(lngIndex)
    Next lngIndex

    strResult = Join(arrLines, vbCrLf)

    Get_Prefix_EachLine_Apostrophe = strResult

End Function

Public Function Get_Log_Prefix_ByType(prefixType As enuWriting_Prefix_Type) As String
    ' Ruft die entsprechende Pr�fix-Funktion basierend auf dem �bergebenen Enum-Wert auf
    
    Select Case prefixType
        Case ptInfo
            Get_Log_Prefix_ByType = Get_Log_Prefix_Info()
        Case ptError
            Get_Log_Prefix_ByType = Get_Log_Prefix_Error()
        Case ptWarning
            Get_Log_Prefix_ByType = Get_Log_Prefix_Warning()
        Case Else
            Get_Log_Prefix_ByType = "[UNKNOWN] "
    End Select
End Function
Public Function Get_Log_Prefix_Info() As String
    ' Gibt ein Prefix f�r Info-Nachrichten zur�ck
    Get_Log_Prefix_Info = "[INFO] "
End Function

Public Function Get_Log_Prefix_Warning() As String
    ' Gibt ein Prefix f�r Warnungen zur�ck
    Get_Log_Prefix_Warning = "[WARNING] "
End Function

Public Function Get_Log_Prefix_Error() As String
    ' Gibt ein Prefix f�r Fehlermeldungen zur�ck
    Get_Log_Prefix_Error = "[ERROR] "
End Function
'############################################ Message #######################################################
Public Function Get_Message_Welcome_User(strUserName As String) As String
    ' Erstellt eine Willkommensnachricht f�r den Benutzer
    Get_Message_Welcome_User = "Willkommen, " & strUserName & "!" & vbCrLf & _
                        "Heute ist " & Format(Date, "dddd, dd. mmmm yyyy") & "." & vbCrLf
End Function
Public Function Get_Message_Form_Opened(strFormName As String) As String

    Get_Message_Form_Opened = "Formular " & strFormName & " wurde ge�ffnet."

End Function


Public Function FormatArrayAsTable(varArray As Variant) As String

    ' Formatiert ein 1D- oder 2D-Array als Text-Tabelle und gibt den gesamten Text als String zur�ck

    
    Dim strResult As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngRows As Long
    Dim lngCols As Long
    Dim strValue As String

    On Error GoTo Handle1D
    lngRows = UBound(varArray, 1)
    lngCols = UBound(varArray, 2)

    ' 2D-Array
    For lngRow = 1 To lngRows
        For lngCol = 1 To lngCols
            strValue = CStr(varArray(lngRow, lngCol))
            strResult = strResult & PadString(strValue, 20)
        Next lngCol
        strResult = strResult & vbCrLf
    Next lngRow

    FormatArrayAsTable = strResult
    Exit Function

Handle1D:
    ' 1D-Array
    On Error GoTo 0
    For lngRow = LBound(varArray) To UBound(varArray)
        strResult = strResult & CStr(varArray(lngRow)) & vbCrLf
    Next lngRow

    FormatArrayAsTable = strResult

End Function


Private Function PadString(strText As String, intWidth As Integer) As String

    ' F�llt einen String rechts mit Leerzeichen auf feste Breite

    
    PadString = Left(strText & Space(intWidth), intWidth)

End Function