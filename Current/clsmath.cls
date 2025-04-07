Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function CalculatePercentage(lngItemsCount As Long, lngItemsCount_Total As Long) As Integer
    ' Berechnet den Fortschritt in Prozent
    
    If lngItemsCount_Total > 0 Then
        CalculatePercentage = (lngItemsCount * 100) \ lngItemsCount_Total
    Else
        CalculatePercentage = 0
    End If
    
End Function