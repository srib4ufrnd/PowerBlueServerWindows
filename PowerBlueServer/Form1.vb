Imports InTheHand.Net.Sockets
Imports InTheHand.Net.Bluetooth
Imports System.IO
Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.ComponentModel



Public Class PowerBlueServerApp



    Dim pptFileName As String
    Dim blueToothClient As New BluetoothClient
    Dim streamRecievedFromBtClient As Net.Sockets.NetworkStream
    Dim pptAppObj As PowerPoint.Application
    Dim presentation As PowerPoint.Presentation
    Dim MyServiceUuid As Guid
    Dim lsnr As BluetoothListener



    Private Sub BrowsePptButton_Click(sender As Object, e As EventArgs) Handles BrowsePptButton.Click

        'Clears all the logs
        PowerBlueLogTextBox.Clear()

        'Call the system requirements message box.
        showTheImportantNoteAndSysRequirementsMessage()

        Dim answer As MsgBoxResult
        answer = MsgBox("All the current running Power point applications will be closed. Are you OK?", MsgBoxStyle.YesNo)
        If answer = MsgBoxResult.Yes Then

            closeAlreadyOpenedPowerPointApp()

            pptFileName = Nothing

            If PptFileSelectionDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
                pptFileName = PptFileSelectionDialog.FileName
            End If

            If (pptFileName IsNot Nothing) Then

                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "PPT File Selected: " & pptFileName & vbCrLf & vbCrLf)

                If pptFileName.EndsWith(".pptx") Then
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "The Seclected File is Valid Presentation: .PPTX." & vbCrLf & vbCrLf)
                    StartServerButton.Enabled = True
                    'MsgBox("The Selected file is PPTX")
                ElseIf pptFileName.EndsWith(".ppt") Then
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "The Seclected File is Valid Presentation: .PPT." & vbCrLf & vbCrLf)
                    StartServerButton.Enabled = True
                    'MsgBox("The Selected file is PPT")
                Else
                    StartServerButton.Enabled = False
                    StopServerButton.Enabled = False
                    MsgBox("The Selected File is Neither PPTX nor PPT.")
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "The Selected File is invalid presenation: Neither .PPTX nor .PPT." & vbCrLf & vbCrLf)
                    'MsgBox("The Selected File is Neither PPTX nor PPT.Please select a Valid Presentation file with extension PPT or PPTX")
                End If
            ElseIf (pptFileName Is Nothing) Then
                StartServerButton.Enabled = False
                StopServerButton.Enabled = False
                MsgBox("No Presentation is selecetd. Please Select a valid PPTX or PPT.")
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "No Presentation is selecetd. Please Select a valid PPTX or PPT." & vbCrLf & vbCrLf)
            End If
        ElseIf answer = MsgBoxResult.No Then
            StartServerButton.Enabled = False
            StopServerButton.Enabled = False
            MsgBox("The Server cannot be Started without closing all the current running Power Point applications.")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "The Server cannot be Started without closing all the current running Power Point applications." & vbCrLf & vbCrLf)
        End If

    End Sub

    Private Sub StartServerButton_Click(sender As Object, e As EventArgs) Handles StartServerButton.Click
        closeAlreadyOpenedPowerPointApp()
        minimizePowerBlueServerApp()
        PowerBlueServerBackgroundWorker.RunWorkerAsync()




    End Sub

    Private Sub StopServerButton_Click(sender As Object, e As EventArgs) Handles StopServerButton.Click
        'PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Stopped.")

        Dim answer As MsgBoxResult
        answer = MsgBox("Do you want to stop server And Quit now?", MsgBoxStyle.YesNo)
        If answer = MsgBoxResult.Yes Then

            If ((PowerBlueServerBackgroundWorker.CancellationPending = False) AndAlso (PowerBlueServerBackgroundWorker.IsBusy = True)) Then

                PowerBlueServerBackgroundWorker.Dispose()
                PowerBlueServerBackgroundWorker.CancelAsync()


            End If


            'If True then the server will be closed
            'If False then the server will only be rolled back.
            'rollBackServerToInitialOrClose(True)
            'The below END command will terminate this application completly.
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Major Problem occured.  So terminating the server. Please restart the server." & vbCrLf & vbCrLf)

            closeAlreadyOpenedPowerPointApp()

            MsgBox("Ooops! Terminating the program   :-(" & vbCrLf & vbCrLf & "Possible Reasons:" & vbCrLf & vbCrLf & "1. User Stoped the server. Force close." & vbCrLf & vbCrLf & "2. Client disconnected from the server." & vbCrLf & vbCrLf & "3. Power Point Closed. No Power Point Object To Operate." & vbCrLf & vbCrLf & "4. User forcely interrupted/closed the PowerPoint/SlideShow." & vbCrLf & vbCrLf & "5. Major Problem occured. So Force closing the server.")
            'Completly closes the application.
            End

        End If


    End Sub

    Private Sub minimizePowerBlueServerApp()
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub closeAlreadyOpenedPowerPointApp()


        'Need to get the COM object for currently running power point application.
        'pptAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")

        'GetObject will get new PowerPoint Application COM object if no PPT is opened.
        pptAppObj = GetObject("", "PowerPoint.Application")


        If Not TypeName(pptAppObj) = "Empty" Then
            'MsgBox("This closes the current running Power point application to Start the Server.")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Closing the current running Power point application to Start the Server." & vbCrLf & vbCrLf)
            exitPowerPoint()
        Else

        End If


    End Sub

    Private Function isPowerPointAppRunning() As Boolean
        'Need to get the COM object for currently running power point application.
        'pptAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")

        'GetObject will get new PowerPoint Application COM object if no PPT is opened.
        pptAppObj = GetObject("", "PowerPoint.Application")

        If Not TypeName(pptAppObj) = "Empty" Then
            MsgBox("Already Powerpoint application is Running. Please close it to Start the Server.")
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub ClearLogsButton_Click(sender As Object, e As EventArgs) Handles ClearLogsButton.Click
        PowerBlueLogTextBox.Clear()
    End Sub

    Private Sub PowerBlueServerBackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles PowerBlueServerBackgroundWorker.DoWork
        'MsgBox("Power Blue Server Started...")

        ' Do not access the form's BackgroundWorker reference directly. 
        ' Instead, use the reference provided by the sender parameter. 
        'Dim bw As BackgroundWorker = CType(sender, BackgroundWorker)

        ' Extract the argument. 
        'Dim arg As Integer = Fix(e.Argument)

        ' Start the time-consuming operation.
        'e.Result = TimeConsumingOperation(bw, arg)
        'e.Result = startDummyBluetoothServerInANewThread(bw)
        'e.Result = startBluetoothServerInANewThread(bw)

        'startDummyBluetoothServerInANewThread()
        startBluetoothServerInANewThread()

        ' If the operation was canceled by the user,  
        ' set the DoWorkEventArgs.Cancel property to true. 
        'If bw.CancellationPending Then
        'e.Cancel = True
        'End If

    End Sub

    Private Sub PowerBlueServerBackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles PowerBlueServerBackgroundWorker.RunWorkerCompleted
        'MsgBox("Power Blue Server Stopped.")

        If (e.Error IsNot Nothing) Then
            ' There was an error during the operation. 
            Dim msg As String = String.Format("An error occurred: {0}", e.Error.Message)
            'MessageBox.Show(msg)
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Blue Server Stopped." & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)

        ElseIf e.Cancelled Then
            ' Next, handle the case where the user canceled the  
            ' operation. 
            ' Note that due to a race condition in  
            ' the DoWork event handler, the Cancelled 
            ' flag may not have been set, even though 
            ' CancelAsync was called.
            ' The user canceled the operation.
            'MessageBox.Show("Operation was canceled")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Blue Server Stopped." & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        Else
            ' Finally, handle the case where the operation succeeded.
            ' The operation completed normally. 
            Dim msg As String = String.Format("Result = {0}", e.Result)
            'MessageBox.Show(msg)
        End If




    End Sub



    Private Sub startDummyBluetoothServerInANewThread()


        PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Started...")
        StartServerButton.Enabled = False
        BrowsePptButton.Enabled = False
        StopServerButton.Enabled = True


        Thread.Sleep(3000)
        controlPowerPointWithCommand("Open")
        'Thread.Sleep(3000)
        'MsgBox("Total Slides = " & getNumberOfSlidesInPresentation())
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Strt")
        'Thread.Sleep(3000)
        'MsgBox("Current Slide = " & getCurrentSlidePosition())
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Whit")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Norm")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Next")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Blac")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Norm")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Norm")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Norm")
        'Thread.Sleep(3000)
        'MsgBox("Total Slides = " & getCurrentSlidePosition())
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Next")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Prev")
        'Thread.Sleep(3000)
        'MsgBox("Total Slides = " & getCurrentSlidePosition())
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Next")
        'Thread.Sleep(3000)
        'MsgBox("Total Slides = " & getCurrentSlidePosition())
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Prev")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Prev")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Last")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Frst")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Last")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Frst")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("G001")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("G001")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("G002")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("G002")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("G003")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("G005")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Stop")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Exit")


        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Strt")

        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Stop")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Exit")


        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Strt")

        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Exit")


        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")

        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Exit")


        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")

        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Exit")


        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")

        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")


        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")

        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Open")


    End Sub




    Private Sub startBluetoothServerInANewThread()

        'MsgBox("startBluetoothServerInANewThread: Inside Method")


        PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Blue Server Started..." & vbCrLf & vbCrLf)
        StartServerButton.Enabled = False
        BrowsePptButton.Enabled = False
        StopServerButton.Enabled = True



        'Dim received(1024) As Byte
        Dim received(3) As Byte
        'received = Nothing




        MyServiceUuid = New Guid("{94f39d29-7d6d-437d-973b-fba39e49d4ee}")

        'MyServiceUuid = Guid.NewGuid()

        'MsgBox(MyServiceUuid.ToString)

        lsnr = New BluetoothListener(MyServiceUuid)
        lsnr.Start()
        'MsgBox("Server is Listening")
        'MsgBox("Server started, waiting for clients")

        blueToothClient = lsnr.AcceptBluetoothClient()

        Try

            If (blueToothClient IsNot Nothing AndAlso blueToothClient.Connected) Then
                'MsgBox("startBluetoothServerInANewThread: Client: " & blueToothClient.RemoteMachineName & " :Connected")
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Client: " & blueToothClient.RemoteMachineName & " :Connected" & vbCrLf & vbCrLf)

                streamRecievedFromBtClient = blueToothClient.GetStream()
                If (streamRecievedFromBtClient IsNot Nothing) Then
                    'MsgBox("startBluetoothServerInANewThread: Inside If Loop: Begin Read stream")
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Listening for commands from Client" & vbCrLf & vbCrLf)
                    streamRecievedFromBtClient.BeginRead(received, 0, received.Length, New AsyncCallback(AddressOf ReadCallBack), received)
                End If
            Else
                'MsgBox("startBluetoothServerInANewThread: Inside Else Loop: blueToothClient is nothing")
                rollBackServerToInitialOrClose(True)

            End If
        Catch ex As Exception
            'MsgBox("startBluetoothServerInANewThread: Inside Catch: Got Exception")
            'MsgBox("Exception Occured" & vbCrLf & ex.Message)
            rollBackServerToInitialOrClose(True)
        End Try



    End Sub


    Private Sub ReadCallBack(ar As IAsyncResult)

        'MsgBox("ReadCallBack: Inside Method")

        Try

            'Dim received(1024) As Byte
            Dim receivedLocal(3) As Byte
            Dim pptControllingCommandFull As String
            Dim pptControllingCommandTrim As String
            Dim pptControllingCommand As String
            pptControllingCommandFull = Nothing
            'received = Nothing


            If (blueToothClient IsNot Nothing AndAlso blueToothClient.Connected) Then

                If ((ar IsNot Nothing) AndAlso ar.IsCompleted) Then

                    receivedLocal = ar.AsyncState
                    pptControllingCommandFull = System.Text.UTF8Encoding.ASCII.GetString(receivedLocal)
                    'MsgBox("ReadCallBack: Inside Method: Asc Encode pptControllingCommand = " & AscEncode(pptControllingCommandFull) & ":")
                    'MsgBox("ReadCallBack: Inside Method: Chr Encode pptControllingCommand = " & ChrEncode(pptControllingCommandFull) & ":")
                    'The pptControllingCommandFull is coming as string with 1024 bytes. SO i need to trim the string to get the valid string length

                    If (Not (String.IsNullOrEmpty(pptControllingCommandFull)) AndAlso (pptControllingCommandFull.Length() = 4) AndAlso Not (AscEncode(pptControllingCommandFull).StartsWith("0000")) AndAlso Not (AscEncode(pptControllingCommandFull).EndsWith("0000"))) Then
                        'MsgBox("ReadCallBack: Inside Method: pptControllingCommandFull = " & pptControllingCommandFull & ":")
                        pptControllingCommandTrim = pptControllingCommandFull.Trim()
                        'MsgBox("ReadCallBack: Inside Method: pptControllingCommandFullTrim = " & pptControllingCommandTrim & ":")
                        pptControllingCommand = pptControllingCommandTrim.Replace("\xA0", String.Empty)
                        'MsgBox("ReadCallBack: Inside Method: pptControllingCommand = " & pptControllingCommand & ":")
                        controlPowerPointWithCommand(pptControllingCommand)



                        If (blueToothClient IsNot Nothing AndAlso blueToothClient.Connected) Then

                            If (streamRecievedFromBtClient IsNot Nothing) Then

                                streamRecievedFromBtClient.Flush()
                                'received = Nothing
                                'MsgBox("ReadCallBack: Inside If Loop: Begin Read stream")
                                Dim receivedLocalAnother(3) As Byte
                                streamRecievedFromBtClient.BeginRead(receivedLocalAnother, 0, receivedLocalAnother.Length, New AsyncCallback(AddressOf ReadCallBack), receivedLocalAnother)
                            Else
                                'MsgBox("ReadCallBack: Inside Else Loop0: blueToothClient is nothing To Recurse")
                                'rollBackServerToInitialOrClose(False)
                                rollBackServerToInitialOrClose(True)

                            End If
                        Else
                            'MsgBox("ReadCallBack: Inside Else Loop1: blueToothClient is nothing To Recurse")
                            'rollBackServerToInitialOrClose(False)
                            rollBackServerToInitialOrClose(True)
                        End If
                    Else
                        'MsgBox("ReadCallBack: Inside Else Loop2: blueToothClient is nothing To Recurse")
                        'rollBackServerToInitialOrClose(False)
                        rollBackServerToInitialOrClose(True)

                    End If
                Else
                    'MsgBox("ReadCallBack: Inside Else Loop1: blueToothClient is nothing")
                    'rollBackServerToInitialOrClose(False)
                    rollBackServerToInitialOrClose(True)

                End If
            Else
                'MsgBox("ReadCallBack: Inside Else Loop2: blueToothClient is nothing")
                'rollBackServerToInitialOrClose(False)
                rollBackServerToInitialOrClose(True)
            End If

        Catch ex As Exception
            'MsgBox("ReadCallBack: Inside Catch: Got Exception")
            'MsgBox("Exception Occured" & vbCrLf & ex.Message)
            rollBackServerToInitialOrClose(True)

        End Try


    End Sub


    Private Sub controlPowerPointWithCommand(pptControllingCommand As String)
        'MsgBox("Command:" & pptControllingCommand & ":")
        PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: " & pptControllingCommand & ":" & vbCrLf & vbCrLf)

        If ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Open")) AndAlso (pptControllingCommand.EndsWith("Open"))) Then
            'MsgBox("Command Recieved: Open")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Open" & vbCrLf & vbCrLf)
            'Do the coding here
            openPowerPoint(pptFileName)
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Exit")) AndAlso (pptControllingCommand.EndsWith("Exit"))) Then
            'MsgBox("Command Recieved: Exit")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Exit" & vbCrLf & vbCrLf)
            closeAlreadyOpenedPowerPointApp()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Strt")) AndAlso (pptControllingCommand.EndsWith("Strt"))) Then
            'MsgBox("Command Recieved: Strt")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Strt" & vbCrLf & vbCrLf)
            'Do the coding here
            startPowerPointSlideShow()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Stop")) AndAlso (pptControllingCommand.EndsWith("Stop"))) Then
            'MsgBox("Command Recieved: Stop")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Stop" & vbCrLf & vbCrLf)
            'Do the coding here
            stopPowerPointSlideShow()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Rsrt")) AndAlso (pptControllingCommand.EndsWith("Rsrt"))) Then
            'MsgBox("Command Recieved: Rsrt")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Rsrt" & vbCrLf & vbCrLf)
            'Do the coding here
            restartPowerPointFromFirstSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Prev")) AndAlso (pptControllingCommand.EndsWith("Prev"))) Then
            'MsgBox("Command Recieved: Prev")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Prev" & vbCrLf & vbCrLf)
            movePowerPointToPreviousSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Next")) AndAlso (pptControllingCommand.EndsWith("Next"))) Then
            'MsgBox("Command Recieved: Next")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Next" & vbCrLf & vbCrLf)
            movePowerPointToNextSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Frst")) AndAlso (pptControllingCommand.EndsWith("Frst"))) Then
            'MsgBox("Command Recieved: Frst")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Frst" & vbCrLf & vbCrLf)
            movePowerPointToFirstSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Last")) AndAlso (pptControllingCommand.EndsWith("Last"))) Then
            'MsgBox("Command Recieved: Last")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Last" & vbCrLf & vbCrLf)
            movePowerPointToLastSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Whit")) AndAlso (pptControllingCommand.EndsWith("Whit"))) Then
            'MsgBox("Command Recieved: Whit")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Whit" & vbCrLf & vbCrLf)
            'Do the coding here
            displayWhiteBackGround()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Blac")) AndAlso (pptControllingCommand.EndsWith("Blac"))) Then
            'MsgBox("Command Recieved: Blac")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Blac" & vbCrLf & vbCrLf)
            'Do the coding here
            displayBlackBackGround()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Norm")) AndAlso (pptControllingCommand.EndsWith("Norm"))) Then
            'MsgBox("Command Recieved: Norm")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As: Norm" & vbCrLf & vbCrLf)
            'Do the coding here
            displayRunningBackGround()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("G"))) Then
            'MsgBox("Command Recieved: Blac")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Processing The Command From Client As:" & pptControllingCommand & vbCrLf & vbCrLf)
            'Do the coding here
            goToTheSlideNum(pptControllingCommand)
        End If


    End Sub


    Private Sub openPowerPoint(pptFileName As String)
        Try
            If (pptAppObj Is Nothing) Then
                'pptAppObj = CreateObject("powerpoint.application")
                closeAlreadyOpenedPowerPointApp()
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Successfully Closed All PowerPoint Applications.")
                PowerBlueLogTextBox.AppendText(vbCrLf & "In the Process Of Creating PPT Application.....")
                pptAppObj = New PowerPoint.Application
                PowerBlueLogTextBox.AppendText(vbCrLf & "Created New Power Point Application Object To Work On.")
                If (presentation Is Nothing) Then
                    PowerBlueLogTextBox.AppendText(vbCrLf & "Both PowerPoint And Presentation Are Nothing.")
                    PowerBlueLogTextBox.AppendText(vbCrLf & "Creating And Preparing The Presentation To Open.")
                    presentation = pptAppObj.Presentations.Open(FileName:=pptFileName, ReadOnly:=True)
                    PowerBlueLogTextBox.AppendText(vbCrLf & "Succesfully Created And Opened New Presentation.")
                    'pptAppObj.Visible = True
                End If

            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To close")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To close" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub exitPowerPoint()

        Try
            If (presentation IsNot Nothing) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                presentation.Close()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(presentation)
                presentation = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If

            If (pptAppObj IsNot Nothing) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
                pptAppObj.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pptAppObj)
                pptAppObj = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If

        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To close")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To close" & vbCrLf & vbCrLf)
            'rollBackServerToInitialOrClose(True)
            'The below END command will terminate this application completly.
            MsgBox("Ooops! Terminating the program   :-(" & vbCrLf & vbCrLf & "Possible Reasons:" & vbCrLf & vbCrLf & "1. User Stoped the server. Force close." & vbCrLf & vbCrLf & "2. Client disconnected from the server." & vbCrLf & vbCrLf & "3. Power Point Closed. No Power Point Object To Operate." & vbCrLf & vbCrLf & "4. User forcely interrupted/closed the PowerPoint/SlideShow." & vbCrLf & vbCrLf & "5. Major Problem occured. So Force closing the server.")
            End
        End Try




    End Sub

    Private Sub startPowerPointSlideShow()
        Try
            If (presentation IsNot Nothing) Then
                presentation.SlideShowSettings.Run()
                logCurrentSlideNumberShown()
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)

        End Try
    End Sub

    Private Sub stopPowerPointSlideShow()
        Try
            If (presentation IsNot Nothing) Then
                logCurrentSlideNumberShown()
                presentation.SlideShowWindow.View.Exit()
                'presentation.Close()
                'pptAppObj.Quit()

            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub restartPowerPointFromFirstSlide()
        Try
            If (presentation IsNot Nothing) Then
                presentation.SlideShowWindow.View.First()
                logCurrentSlideNumberShown()
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub movePowerPointToPreviousSlide()
        Try
            If (presentation IsNot Nothing) Then
                If (getCurrentSlidePosition() > 1) Then
                    presentation.SlideShowWindow.View.Previous()
                    logCurrentSlideNumberShown()
                Else
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Cannot Move to Previous Slide. You are currently viewing the First Slide Num : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
                End If
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub movePowerPointToNextSlide()
        Try
            If (presentation IsNot Nothing) Then
                If (getCurrentSlidePosition() < getNumberOfSlidesInPresentation()) Then
                    presentation.SlideShowWindow.View.Next()
                    logCurrentSlideNumberShown()
                Else
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Cannot Move to Next Slide. You are currently viewing the Last Slide Num : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
                End If
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub movePowerPointToFirstSlide()
        Try
            If (presentation IsNot Nothing) Then
                If (getCurrentSlidePosition() > 1) Then
                    presentation.SlideShowWindow.View.First()
                    logCurrentSlideNumberShown()
                Else
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "You are currently viewing the First Slide Num Only : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
                End If

            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub movePowerPointToLastSlide()
        Try
            If (presentation IsNot Nothing) Then
                If (getCurrentSlidePosition() < getNumberOfSlidesInPresentation()) Then
                    presentation.SlideShowWindow.View.Last()
                    logCurrentSlideNumberShown()
                Else
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "You are currently viewing the Last Slide Num Only: " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
                End If

            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub movePowerPointToSlideNum(slideNumToGoIntLocal As Integer)
        Try
            If (presentation IsNot Nothing) Then
                presentation.SlideShowWindow.View.GotoSlide(slideNumToGoIntLocal)
                logCurrentSlideNumberShown()
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub displayWhiteBackGround()
        Try
            If (presentation IsNot Nothing) Then
                presentation.SlideShowWindow.View.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen
                logCurrentSlideNumberShown()
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try

    End Sub

    Private Sub displayBlackBackGround()
        Try
            If (presentation IsNot Nothing) Then
                presentation.SlideShowWindow.View.State = PowerPoint.PpSlideShowState.ppSlideShowBlackScreen
                logCurrentSlideNumberShown()
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub displayRunningBackGround()
        Try
            If (presentation IsNot Nothing) Then
                movePowerPointToSlideNum(getCurrentSlidePosition())
                logCurrentSlideNumberShown()
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub



    Private Sub goToTheSlideNum(pptControllingCommandLocal As String)
        Try
            Dim slideNumToGoStr As String
            Dim slideNumToGoInt As Integer


            slideNumToGoStr = pptControllingCommandLocal.Substring(1, 3)
            slideNumToGoInt = CInt(slideNumToGoStr)

            If (presentation IsNot Nothing) Then
                If (slideNumToGoInt = 1) Then
                    movePowerPointToFirstSlide()
                ElseIf (slideNumToGoInt = getNumberOfSlidesInPresentation()) Then
                    movePowerPointToLastSlide()
                ElseIf (slideNumToGoInt > getNumberOfSlidesInPresentation()) Then
                    PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "The slide Number To Go is > Total Number of slides. So Moving to Last Slide: " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
                    movePowerPointToLastSlide()
                Else
                    movePowerPointToSlideNum(slideNumToGoInt)
                End If
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try

    End Sub

    Private Function getNumberOfSlidesInPresentation() As Integer
        Try
            If (presentation IsNot Nothing) Then
                Return presentation.Slides.Count()
            Else
                Return 0
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Function

    Private Function getCurrentSlidePosition() As Integer
        Try
            If (presentation IsNot Nothing) Then
                Return presentation.SlideShowWindow.View.CurrentShowPosition
            Else
                Return 0
            End If
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Function

    Private Sub logCurrentSlideNumberShown()
        Try
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Current Slide No Showing : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
        Catch ex As Exception
            ' Show the exception's message.
            'MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
            rollBackServerToInitialOrClose(True)
        End Try
    End Sub

    Private Sub rollBackServerToInitialOrClose(Mode As Boolean)
        Try
            'MsgBox("Major Problem occured. So restarting the server to original mode")
            'PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Major Problem occured. So restarting the server to original mode" & vbCrLf & vbCrLf)
            'MsgBox("Major Problem occured. So terminating the server. Please restart the server.")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Major Problem occured.  So terminating the server. Please restart the server." & vbCrLf & vbCrLf)

            closeAlreadyOpenedPowerPointApp()

            If (lsnr IsNot Nothing) Then
                'lsnr.Stop()
                lsnr = Nothing
            End If



            If (blueToothClient IsNot Nothing) Then
                blueToothClient.Close()
                blueToothClient = Nothing
            End If


            If (streamRecievedFromBtClient IsNot Nothing) Then
                streamRecievedFromBtClient.Close()
                streamRecievedFromBtClient = Nothing
            End If


        Catch ex As Exception
            ' Show the exception's message.
            MessageBox.Show(ex.Message)

            ' Show the stack trace, which is a list of methods 
            ' that are currently executing.
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)

            'MsgBox("Power Point Already Stopped. No Power Point Object To Operate")
            'PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Power Point Already Stopped. No Power Point Object To Operate" & vbCrLf & vbCrLf)
        Finally
            lsnr = Nothing
            MyServiceUuid = Nothing
            blueToothClient = Nothing
            streamRecievedFromBtClient = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()

            StartServerButton.Enabled = False
            StopServerButton.Enabled = False
            BrowsePptButton.Enabled = True
            If (Mode = True) Then

                'The below END command will terminate this application completly.
                MsgBox("Ooops! Terminating the program   :-(" & vbCrLf & vbCrLf & "Possible Reasons:" & vbCrLf & vbCrLf & "1. User Stoped the server. Force close." & vbCrLf & vbCrLf & "2. Client disconnected from the server." & vbCrLf & vbCrLf & "3. Power Point Closed. No Power Point Object To Operate." & vbCrLf & vbCrLf & "4. User forcely interrupted/closed the PowerPoint/SlideShow." & vbCrLf & vbCrLf & "5. Major Problem occured. So Force closing the server.")
                End
            ElseIf (Mode = False) Then
                Dim answer As MsgBoxResult

                answer = MsgBox("Client Got Disconnected. Do you want to quit server ?", MsgBoxStyle.YesNo)
                If answer = MsgBoxResult.Yes Then
                    End
                ElseIf answer = MsgBoxResult.No Then
                    MsgBox("The server will be restarted to original mode")
                    StartServerButton.Enabled = False
                    StopServerButton.Enabled = False
                    BrowsePptButton.Enabled = True
                End If
            End If

        End Try
    End Sub


    Function AscEncode(str)
        Dim i
        Dim sAscii

        sAscii = ""
        For i = 1 To Len(str)
            sAscii = sAscii + CStr(Hex(Asc(Mid(str, i, 1))))
            'MsgBox("AscEncode: Inside Method: str =" & str & "Mid(str, i, 1)" & Mid(str, i, 1) & "Asc(Mid(str, i, 1))" & Asc(Mid(str, i, 1)) & "Hex(Asc(Mid(str, i, 1)))" & Hex(Asc(Mid(str, i, 1))) & "CStr(Hex(Asc(Mid(str, i, 1))))" & CStr(Hex(Asc(Mid(str, i, 1)))) & "sAscii" & sAscii)
        Next

        AscEncode = sAscii
    End Function


    Function ChrEncode(str)
        Dim i
        Dim sStr

        sStr = ""
        For i = 1 To Len(str) Step 2
            sStr = sStr + Chr(CLng("&H" & Mid(str, i, 2)))
        Next

        ChrEncode = sStr
    End Function


    Private Sub ServerHelp_Click(sender As Object, e As EventArgs) Handles ServerHelp.Click
        MsgBox("                                                  Instructions To Run Server                                                 " & vbCrLf & vbCrLf & "1.Make sure that all the power point slideshows & Power Point applications are closed before starting Power Blue Server." & vbCrLf & vbCrLf & "2. Make sure that the bluetooth adapter is available in your machine." & vbCrLf & vbCrLf & "3. Make sure that the bluetooth is turned on in your machine." & vbCrLf & vbCrLf & "4. Make sure that the bluetooth adapter is turned on in your machine via Bluetooth Settings." & vbCrLf & vbCrLf & "5. Make sure your machine is made visible to other devices via bluetooth settings." & vbCrLf & vbCrLf & "6. Make sure your machine allows bluetooth devices to connect via bluetooth settings." & vbCrLf & vbCrLf & "7. [Optional] In Case of any strange errors, installing 32feet.net might solve." & vbCrLf & "    Please find download link below." & vbCrLf & "    http://32feet.codeplex.com/releases/view/88941 " & vbCrLf & vbCrLf & "8. For more info please refer to the installation instructions in github link below." & vbCrLf & "    http://bit.ly/1zkKX29" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "  Used 32Feet.net support for developing this application." & vbCrLf & "  Thanks to In The Hand Ltd.  ")
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles ServerImpNote.Click
        showTheImportantNoteAndSysRequirementsMessage()
    End Sub


    Private Sub showTheImportantNoteAndSysRequirementsMessage()

        MsgBox("                                                  VERY VERY IMPORTANT NOTE                                                 " & vbCrLf & vbCrLf & "Power Blue Server Tested well and working well on the below system environment." & vbCrLf & vbCrLf & "Power Blue Server Might or Might not work on all other environments." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Please feel free to install and test the power blue server on other environments as well." & vbCrLf & vbCrLf & "Please inform us at the below github link whether the power blue server is working fine or not in the other environments." & vbCrLf & vbCrLf & "Please do mention the environment details as well in the github link below, with the status as WORKING or NOT_WORKING." & vbCrLf & vbCrLf & "GITHUB LINK: http://bit.ly/1zkKX29" & vbCrLf & vbCrLf & "Thank You." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                  SYSTEM REQUIREMENTS                                                 " & vbCrLf & vbCrLf & "Microsoft Office: Microsoft Office Professional Plus 2010" & vbCrLf & vbCrLf & "Microsoft PowerPoint: Microsoft PowerPoint 2010 - 32 Bit" & vbCrLf & vbCrLf & "Power Point Extensions: PPT & PPTX" & vbCrLf & vbCrLf & "Operating System:  Windows 7 Enterprise - Service Pack 1" & vbCrLf & vbCrLf & "Processor: Intel(R) Core(TM) i7-3635 QM CPU @ 2.40GHz 2.40GHz" & vbCrLf & vbCrLf & "RAM: 8 GB" & vbCrLf & vbCrLf & "System Type: 64-bit Operating System.")

    End Sub

End Class
