Imports System.ComponentModel
Imports System.Threading

Friend Class mainForm
    ' checked: v2.1
    ' introduced multiple files selection with v2.1
    Private rcOpenFileDialog As OpenFileDialog = New OpenFileDialog
    Private theFileNames As List(Of String) = New List(Of String)
    Private theBase As rcBase = New rcBase(Me)
    Private Delegate Sub setFormItemCallback(ByVal which As Label, ByVal aTxt As String)

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        Text = progName & " v" & progVersion_txt & " (" & progLastDate & ")"
        Refresh()
    End Sub

    Private Sub changeFormItemDelegate(ByVal which As Label, ByVal aTxt As String)
        ' checked: v2.0
        ' multi-threading handling according to VisualBasic help, used in modified way

        which.Text = aTxt
        Refresh()
    End Sub

    Friend Sub updateFormItem(ByVal which As Label, ByVal aTxt As String)
        ' checked: v2.0
        ' this sub allows a thread safe change of the form items.
        ' multi-threading handling according to VisualBasic help, used in modified way
        Dim ok As Boolean = True

        If which.InvokeRequired Then
            ' It's on a different thread, so use Invoke.
            Dim d As New setFormItemCallback(AddressOf changeFormItemDelegate)
            Invoke(d, New Object() {[which], [aTxt]})
        Else
            ' It's on the same thread, no need for Invoke.
            changeFormItemDelegate(which, aTxt)
        End If
    End Sub

    Friend Sub updateStatus(Optional ByVal txt As String = "", Optional ByVal isOK As Boolean = True)
        ' checked: v2.0

        Dim presetStr As String = ""

        If isOK Then
            presetStr = "OK"
        Else
            presetStr = "ERROR"
        End If

        If txt = "" Then
            txt = presetStr
        End If

        updateFormItem(statusTxt, txt)
    End Sub

    Private Sub analyzeButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles analyzeButton.Click
        ' checked: v2.1
        analyzeButton.Enabled = False
        aboutButton.Enabled = False
        exitButton.Enabled = False
        cancelWorkButton.Enabled = True
        updateStatus("Analyze data file(s)...")
        progressBar.Value = 0

        rcOpenFileDialog.Multiselect = True

        If rcOpenFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            For Each item As String In rcOpenFileDialog.FileNames
                theFileNames.Add(item)
            Next item

            Try
                rcBackgroundWorker.RunWorkerAsync(workerTypes.ANALYSIS_doRun)
            Catch ex As Exception
                theBase.out.writeLog(ex.Message)
            End Try
            rcOpenFileDialog.FileName = ""
        Else
            analyzeButton.Enabled = True
            aboutButton.Enabled = True
            exitButton.Enabled = True
            cancelWorkButton.Enabled = False
            updateStatus("File selection was not successful.", False)
        End If
    End Sub

    Private Sub cancelWorkButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cancelWorkButton.Click
        ' checked: v2.0
        rcBackgroundWorker.CancelAsync()
        cancelWorkButton.Enabled = False
        updateStatus("The work process will be cancelled as soon as possible.")
    End Sub

    Private Sub aboutButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aboutButton.Click
        ' checked: v2.0
        rcAboutBox.ShowDialog()
    End Sub

    Private Sub exitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles exitButton.Click
        ' checked: v2.0

        exitButton.Enabled = False
        updateStatus("Writing log and result files into your main document folder...")
        analyzeButton.Enabled = False
        cancelWorkButton.Enabled = False
        aboutButton.Enabled = False
        Refresh()

        theBase.cleanup()
        Close()
    End Sub

    Private Sub rcBackgroundWorker_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles rcBackgroundWorker.DoWork
        ' checked: v2.1
        ' This event handler is where the actual work is done.

        ' Get the BackgroundWorker object that raised this event.
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim result As workerTypes = workerTypes.WORKER_Nothing
        Dim dummy As workerTypes = workerTypes.WORKER_Nothing

        ' (ps) changed the common e.Result usage to create a simple signal semaphore
        ' note: e.Result must not be set if e.Cancel is set to True

        Select Case e.Argument
            Case workerTypes.ANALYSIS_doRun
                result = theBase.analyzeFile(worker, e, theFileNames)
        End Select

        If worker.CancellationPending Then
            ' last chance to catch it :-)
            e.Cancel = True
        End If

        If Not e.Cancel Then
            ' this if clause is necessary not to crash the program
            e.Result = result
        End If
    End Sub

    Private Sub rcBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles rcBackgroundWorker.ProgressChanged
        ' checked: v2.0
        progressBar.Value = e.ProgressPercentage
    End Sub

    Private Sub rcBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles rcBackgroundWorker.RunWorkerCompleted
        ' checked: v2.0
        Dim whereFrom As workerTypes = workerTypes.WORKER_Nothing

        If (e.Error IsNot Nothing) Then
            ' First, handle the case where an exception was thrown.
            MessageBox.Show(e.Error.Message)
            theBase.out.writeLog()
            theBase.out.writeLog("*** System Error: ***")
            theBase.out.writeLog(e.Error.Message)
            theBase.out.writeLog()
            exitButton.Enabled = True
        ElseIf e.Cancelled Then
            ' Second, handle a cancelled work process.
            updateStatus("The work process was cancelled.", False)
            theBase.out.writeLog()
            theBase.out.writeLog("***************************************************")
            theBase.out.writeLog("*** The work process was cancelled by the user. ***")
            theBase.out.writeLog("***************************************************")
            theBase.out.writeLog()
            exitButton.Enabled = True
        Else
            ' Finally, handle the case where the operation succeeded.
            whereFrom = e.Result
            Select Case whereFrom
                ' ------ ANALYSIS
                Case workerTypes.ANALYSIS_resultOK
                    ' last, adjust GUI
                    cancelWorkButton.Enabled = False
                    aboutButton.Enabled = True
                    exitButton.Enabled = True
                    updateStatus()

                Case workerTypes.ANALYSIS_resultERROR
                    cancelWorkButton.Enabled = False
                    aboutButton.Enabled = True
                    exitButton.Enabled = True
                    updateStatus("", False)

                Case Else
                    cancelWorkButton.Enabled = False
                    aboutButton.Enabled = True
                    exitButton.Enabled = True
                    updateStatus("Unknown Error.", False)
            End Select
        End If
    End Sub
End Class
