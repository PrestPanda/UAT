VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_000_frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Component As New clsComponent_Old
Dim Log As New clsLog
Dim Coding_Git As New clsCoding_Git


Private Sub cmdExportComponents_Click()

    Dim strCommitMessage As String
    
    strCommitMessage = Me.txtCommitMessage.Value


    Dim msgEmptyCommitMessage As VbMsgBoxResult
    
    If txtCommitMessage.Value = "" Or IsNull(txtCommitMessage) = True Then
        msgEmptyCommitMessage = _
            MsgBox("Möchten Sie den Commit wirklich ohne eine Nachricht durchführen?", vbYesNo)
    End If

    If msgEmptyCommitMessage = vbYes Or IsNull(txtCommitMessage) = False Then
        Component.ExportAllComponents
        Coding_Git.CommitPushCurrentComponents strCommitMessage
    End If
    
    Set Coding_Git = Nothing
    
    DoCmd.OpenForm "000_frmMainPage"

End Sub

'Private Sub cmdFormPopUpAdd_Click()
'
'    Dim strFormName As String
'
'    strFormName = InputBox("Bitte geben Sie einen Namen für das Formular ein.")
'
'    Access_Form_CreateNe strFormName
'    Access_Form_Settings_PopUp strFormName
'
'End Sub

Private Sub cmdOpenApplications_Click()

    DoCmd.OpenForm "210_frmApplications", acNormal

End Sub

Private Sub cmdOpenClassBuilder_Click()

    DoCmd.OpenForm "110_frmClassBuilder", acNormal

End Sub

Private Sub Form_Load()

    Log.Write_Application_WelcomeMessage "UAT"

End Sub
