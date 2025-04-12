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

Dim Log As New clsLog


Private Sub cmdExportComponents_Click()

    Dim strCommitMessage As String
    
    strCommitMessage = Me.txtCommitMessage.value

    Dim Coding_Git As New clsCoding_Git
    Dim msgEmptyCommitMessage As VbMsgBoxResult
    
    If txtCommitMessage.value = "" Or IsNull(txtCommitMessage) = True Then
        msgEmptyCommitMessage = _
            MsgBox("Möchten Sie den Commit wirklich ohne eine Nachricht durchführen?", vbYesNo)
    End If

    If msgEmptyCommitMessage = vbYes Or IsNull(txtCommitMessage) = False Then
        ExportAllComponents
        Coding_Git.CommitPushCurrentComponents strCommitMessage
    End If
    
    Set Coding_Git = Nothing
    
    DoCmd.OpenForm "000_frmMainPage"

End Sub

Private Sub cmdOpenClassBuilder_Click()

    DoCmd.OpenForm "110_frmClassBuilder", acNormal, , , acFormAdd, acWindowNormal

End Sub

Private Sub Form_Load()

    Log.Write_Application_WelcomeMessage "UAT"

End Sub
