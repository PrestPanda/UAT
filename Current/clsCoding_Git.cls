Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const strRepoPath As String = "C:\Users\Prest\Desktop\JCs Ultimate Access Tool\Versions\Current\"

Dim objShell As Object

Public Sub CommitPushCurrentComponents(strCommitMessage As String)

    StagingBranch_AddAll
    Commit "Es hat geklappt."
    Push
    
'    Call RunCmdGit("cd /d """ & strRepoPath & """ && git add .")
'
'    If strCommitMessage = "" Then
'        Call RunCmdGit("cd /d """ & strRepoPath & """ && git commit --allow-empty-message -m """"")
'    Else
'        Call RunCmdGit("cd /d """ & strRepoPath & """ && git commit -m """ & strCommitMessage & """")
'    End If
'
'    Call RunCmdGit("cd /d """ & strRepoPath & """ && git push origin master")

End Sub

Private Sub StagingBranch_AddAll()

    RunCommand "cd /d """ & strRepoPath & """ && git add ."

End Sub
Private Sub Commit(strCommitMessage As String)

     If strCommitMessage = "" Then
        RunCommand "cd /d """ & strRepoPath & """ && git commit --allow-empty-message -m """""
    Else
        RunCommand "cd /d """ & strRepoPath & """ && git commit -m """ & strCommitMessage & """"
    End If

End Sub
Private Sub Push()

    RunCommand "cd /d """ & strRepoPath & """ && git push origin master"

End Sub
Private Sub RunCommand(strFullCommand As String)

    objShell.Run "cmd.exe /c " & strFullCommand, 1, True

End Sub

Private Sub Class_Initialize()

    Set objShell = CreateObject("WScript.Shell")

End Sub

Private Sub Class_Terminate()

    Set objShell = Nothing

End Sub