VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_112_frmCodeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Coding_VBA_Write As New clsCoding_VBA_Write
Dim Coding_VBA_Analyze As New clsCoding_VBA_Analyze

Private Sub cmdGenerateCode_Click()
'Working On

    Dim Properties As Variant
    
    Properties = Coding_VBA_Analyze.Get_CodeElements_AsArray


End Sub
