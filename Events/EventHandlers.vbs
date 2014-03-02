'   Sub OnClientConnect(oClient)
'   End Sub

'   Sub OnSMTPData(oClient, oMessage)
'   End Sub

'   Sub OnAcceptMessage(oClient, oMessage)
'   End Sub

'   Sub OnDeliveryStart(oMessage)
'   End Sub

'   Sub OnDeliverMessage(oMessage)
'   End Sub

'   Sub OnBackupFailed(sReason)
'   End Sub

'   Sub OnBackupCompleted()
'   End Sub

'   Sub OnError(iSeverity, iCode, sSource, sDescription)
'   End Sub

'   Sub OnDeliveryFailed(oMessage, sRecipient, sErrorMessage)
'   End Sub

'   Sub OnExternalAccountDownload(oFetchAccount, oMessage, sRemoteUID)
'   End Sub

Sub ReceiveFromMT4(oMessage)
    Dim objShell
    Dim objFSO
    Dim objTextFile
    Dim filename

    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    filename = objShell.ExpandEnvironmentStrings("%TEMP%") & "\mail-from-mt4.txt"
    Set objTextFile = objFSO.CreateTextFile(filename, True)

    objTextFile.WriteLine(oMessage.Body)
    objTextFile.Close

    Set objTextFile = Nothing
    Set objFSO = Nothing
    Set objShell = Nothing
End Sub