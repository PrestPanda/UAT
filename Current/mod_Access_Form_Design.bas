Option Compare Database
Option Explicit

Public Sub Access_Form_SetDesign_DarkMode(strFormName As String)

    Dim varDesign As Variant
    varDesign = GetDesign_AsArray()

    Access_Form_ApplyDesign strFormName, varDesign

End Sub
Public Sub Access_Form_ApplyDesign(strFormName As String, _
    varDesign As Variant)

    ' Wendet ein Farbschema auf ein Formular und dessen Steuerelemente an.
    
    Dim objForm As Form
    Dim objControl As Control
    Dim lngIndex As Long
    
    DoCmd.OpenForm strFormName, acDesign, , , , acWindowNormal
    Set objForm = Forms(strFormName)

    objForm.Section(0).BackColor = CLng(varDesign(0))
    
    For Each objControl In objForm.Controls
        Select Case objControl.ControlType
        
            Case acTextBox, acComboBox
                objControl.BackColor = varDesign(2)
                objControl.ForeColor = varDesign(3)
                objControl.BorderColor = varDesign(4)
                
            Case acCommandButton
                objControl.BackColor = varDesign(5)
                objControl.ForeColor = varDesign(6)
                objControl.BorderColor = varDesign(4)
                
            Case acLabel
                objControl.ForeColor = varDesign(7)
                
            Case acListBox
                objControl.BackColor = varDesign(8)
                objControl.ForeColor = varDesign(1)
                objControl.BorderColor = varDesign(4)
                
            Case acTabCtl
                ApplyDesignToTabControl strFormName, objControl.Name, varDesign
        End Select
    Next objControl

    DoCmd.Save acForm, strFormName
    DoCmd.Close acForm, strFormName, acSaveNo

End Sub

Private Sub ApplyDesignToTabControl(strFormName As String, _
    strTabName As String, _
    varDesign As Variant)

    ' Wendet das Design auf alle Seiten und Steuerelemente einer Registerkarte an.

    Dim objForm As Form
    Dim objTab As Control
    Dim objPage As Page
    Dim objSubControl As Control

    Set objForm = Forms(strFormName)
    Set objTab = objForm.Controls(strTabName)

    For Each objPage In objTab.Pages


        For Each objSubControl In objPage.Controls

            Select Case objSubControl.ControlType

                Case acTextBox, acComboBox
                    objSubControl.BackColor = varDesign(2)
                    objSubControl.ForeColor = varDesign(3)
                    objSubControl.BorderColor = varDesign(4)

                Case acCommandButton
                    objSubControl.BackColor = varDesign(5)
                    objSubControl.ForeColor = varDesign(6)
                    objSubControl.BorderColor = varDesign(4)

                Case acLabel
                    objSubControl.ForeColor = varDesign(7)

                Case acListBox
                    objSubControl.BackColor = varDesign(8)
                    objSubControl.ForeColor = varDesign(1)
                    objSubControl.BorderColor = varDesign(4)

            End Select

        Next objSubControl

    Next objPage

End Sub
Public Function GetDesign_AsArray() As Variant

    ' Gibt ein Array mit vordefinierten RGB-Farben für ein dunkles Formular-Design zurück.
    
    Dim varDesign(0 To 9) As Variant
    
    varDesign(0) = RGB(45, 45, 48)      ' Hintergrund
    varDesign(1) = RGB(241, 241, 241)   ' Textfarbe
    varDesign(2) = RGB(63, 63, 70)      ' Eingabefeld Hintergrund
    varDesign(3) = RGB(241, 241, 241)   ' Eingabefeld Text
    varDesign(4) = RGB(104, 104, 104)   ' Rahmenfarbe
    varDesign(5) = RGB(140, 140, 140)       ' Button Hintergrund
    varDesign(6) = RGB(241, 241, 241)   ' Button Textfarbe
    varDesign(7) = RGB(153, 153, 153)   ' Label Sekundärtext
    varDesign(8) = RGB(37, 37, 38)      ' Wechselhintergrund (z. B. für Listen)
    varDesign(9) = RGB(62, 62, 66)      ' Linien/Fokus-Rahmen

    GetDesign_AsArray = varDesign

End Function