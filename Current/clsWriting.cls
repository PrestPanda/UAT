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
'Gibt die Standardlänge für bestimmte Objekte an z.B. clsLog, Debug-Fenster

    Select Case CallingObject
    
        Case pboDebug
            Get_ProgressBar_StandardLength = 40
        Case Else
            Get_ProgressBar_StandardLength = 0
            
    End Select

End Function
Private Function Get_Separator_StandardLength(CallingObject As enuWriting_Object) As Long
'Gibt die Standardlänge für bestimmte Objekte an z.B. clsLog, Debug-Fenster

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
    
    lngTotalLength = Get_ProgressBar_StandardLength(CallingObject) ' Länge der Fortschrittsleiste
    lngPercentage = Math.CalculatePercentage(lngItemsCount, lngItemsCount_Total)
    
    ' Berechnet die Anzahl der gefüllten Blöcke
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
' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen in der angegebenen Länge

    Dim lngTotalLength As Long
    
    lngTotalLength = Get_ProgressBar_StandardLength(CallingObject) ' Länge der Fortschrittsleiste
        
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
    ' Erstellt eine Trennlinie mit Bindestrichen (-) in der angegebenen Länge
    
    Get_Separator_Dashes = String(lngTotalLength, "-")
    
End Function

Private Function Get_Separator_Equals(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Gleichheitszeichen (=) in der angegebenen Länge
    
    Get_Separator_Equals = String(lngTotalLength, "=")
    
End Function

Private Function Get_Separator_Stars(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Sternen (*) in der angegebenen Länge
    
    Get_Separator_Stars = String(lngTotalLength, "*")
    
End Function

Private Function Get_Separator_Custom(lngTotalLength As Long, strChar As String) As String
    ' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen in der angegebenen Länge
    
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
    
    ' Text wahlweise in Großbuchstaben umwandeln
    If blnUpperCase Then
        strFormattedText = UCase(strText)
    Else
        strFormattedText = strText
    End If
    
    ' Berechnung der Seitenlängen
    lngTextLength = Len(strFormattedText) + 2 ' +2 für die zusätzlichen Leerzeichen
    lngSideLength = (lngTotalLength - lngTextLength) \ 2
    
    ' Falls die Gesamtlänge nicht exakt aufgeteilt werden kann, eine Seite um 1 Zeichen erweitern
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
    ' Erstellt eine Trennlinie mit Bindestrichen (-) in der angegebenen Länge
    Get_Separator_WithText_Dashes = String(lngTotalLength, "-")
End Function

Private Function Get_Separator_WithText_Equals(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Gleichheitszeichen (=) in der angegebenen Länge
    Get_Separator_WithText_Equals = String(lngTotalLength, "=")
End Function

Private Function Get_Separator_WithText_Stars(lngTotalLength As Long) As String
    ' Erstellt eine Trennlinie mit Sternen (*) in der angegebenen Länge
    Get_Separator_WithText_Stars = String(lngTotalLength, "*")
End Function

Private Function Get_Separator_WithText_Custom(lngTotalLength As Long, strChar As String) As String
    ' Erstellt eine Trennlinie mit einem benutzerdefinierten Zeichen in der angegebenen Länge
    Get_Separator_WithText_Custom = String((lngTotalLength \ Len(strChar)), strChar)
End Function
'############################################ TIMESTAMP #######################################################
Public Function Get_Timestamp() As String
    ' Gibt den aktuellen Zeitstempel zurück
    Get_Timestamp = "[" & Format(Now, "yyyy-mm-dd HH:NN:SS") & "] "
End Function
'############################################ PREFIX #######################################################
Public Function Get_Prefix_EachLine_Apostrophe(strInput As String) As String
    ' Fügt jeder Zeile des übergebenen Strings ein Apostroph voran.
    ' Gibt den modifizierten String zurück, der nahtlos in den VBE eingefügt werden kann.

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
    ' Ruft die entsprechende Präfix-Funktion basierend auf dem übergebenen Enum-Wert auf
    
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
    ' Gibt ein Prefix für Info-Nachrichten zurück
    Get_Log_Prefix_Info = "[INFO] "
End Function

Public Function Get_Log_Prefix_Warning() As String
    ' Gibt ein Prefix für Warnungen zurück
    Get_Log_Prefix_Warning = "[WARNING] "
End Function

Public Function Get_Log_Prefix_Error() As String
    ' Gibt ein Prefix für Fehlermeldungen zurück
    Get_Log_Prefix_Error = "[ERROR] "
End Function
'############################################ Message #######################################################
Public Function Get_Message_Welcome_User(strUserName As String) As String
    ' Erstellt eine Willkommensnachricht für den Benutzer
    Get_Message_Welcome_User = "Willkommen, " & strUserName & "!" & vbCrLf & _
                        "Heute ist " & Format(Date, "dddd, dd. mmmm yyyy") & "." & vbCrLf
End Function