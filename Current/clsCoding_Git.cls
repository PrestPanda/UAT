Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const strRepoPath As String = "C:\Users\Prest\Desktop\JCs Ultimate Access Tool\Versions\Current\"

Public Sub CommitPushCurrentComponents(strCommitMessage As String)

    Call RunCmdGit("cd /d """ & strRepoPath & """ && git add .")

    If strCommitMessage = "" Then
        Call RunCmdGit("cd /d """ & strRepoPath & """ && git commit --allow-empty-message -m """"")
    Else
        Call RunCmdGit("cd /d """ & strRepoPath & """ && git commit -m """ & strCommitMessage & """")
    End If

    Call RunCmdGit("cd /d """ & strRepoPath & """ && git push origin master")

End Sub

Private Sub RunCmdGit(strFullCommand As String)

    ' Führt einen Git-Befehl direkt im cmd-Fenster aus und wartet auf die Fertigstellung

    
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")

    objShell.Run "cmd.exe /c " & strFullCommand, 1, True

End Sub