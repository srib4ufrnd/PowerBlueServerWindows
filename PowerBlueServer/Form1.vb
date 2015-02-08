﻿Imports InTheHand.Net.Sockets
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



    Private Sub BrowsePptButton_Click(sender As Object, e As EventArgs) Handles BrowsePptButton.Click

        pptFileName = Nothing

        'MsgBox("My First VB Program")
        If PptFileSelectionDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            pptFileName = PptFileSelectionDialog.FileName
        End If

        If (pptFileName IsNot Nothing) Then

            'PptSelectedLabel.Text = PptSelectedLabel.Text + pptFileName
            PowerBlueLogTextBox.AppendText(vbCrLf & "PPT File Selected: " & pptFileName)

            If pptFileName.EndsWith(".pptx") Then
                'pptStatusLabel.Text = "The Selected file is PPTX"
                PowerBlueLogTextBox.AppendText(vbCrLf & "The Seclected File is Valid Presentation: .PPTX.")
                StartServerButton.Enabled = True
                'MsgBox("The Selected file is PPTX")
            ElseIf pptFileName.EndsWith(".ppt") Then
                'pptStatusLabel.Text = "The Selected file is PPT"
                PowerBlueLogTextBox.AppendText(vbCrLf & "The Seclected File is Valid Presentation: .PPT.")
                StartServerButton.Enabled = True
                'MsgBox("The Selected file is PPT")
            Else
                'pptStatusLabel.Text = "The Selected File is Neither PPTX nor PPT"
                'pptStatusLabel.Text = "The Selected File is Neither PPTX nor PPT.Please select a Valid Presentation file with extension PPT or PPTX"
                StartServerButton.Enabled = False
                StopServerButton.Enabled = False
                MsgBox("The Selected File is Neither PPTX nor PPT.")
                PowerBlueLogTextBox.AppendText(vbCrLf & "The Selected File is invalid presenation: Neither .PPTX nor .PPT.")
                'MsgBox("The Selected File is Neither PPTX nor PPT.Please select a Valid Presentation file with extension PPT or PPTX")
            End If
        ElseIf (pptFileName Is Nothing) Then
            StartServerButton.Enabled = False
            StopServerButton.Enabled = False
            MsgBox("No Presentation is selecetd. Please Select a valid PPTX nor PPT.")
            PowerBlueLogTextBox.AppendText(vbCrLf & "No Presentation is selecetd. Please Select a valid PPTX nor PPT.")
        End If

    End Sub

    Private Sub StartServerButton_Click(sender As Object, e As EventArgs) Handles StartServerButton.Click
        'PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Started...")
        minimizePowerBlueServerApp()
        'closeAlreadyOpenedPowerPointApp()
        'If (isPowerPointAppRunning() = True) Then
        'MsgBox("Already Powerpoint application is Running. Please close it to Start the Server.")
        'Else
        PowerBlueServerBackgroundWorker.RunWorkerAsync()
        'End If




    End Sub

    Private Sub StopServerButton_Click(sender As Object, e As EventArgs) Handles StopServerButton.Click
        'PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Stopped.")
        If ((PowerBlueServerBackgroundWorker.CancellationPending = False) AndAlso (PowerBlueServerBackgroundWorker.IsBusy = True)) Then

            PowerBlueServerBackgroundWorker.Dispose()
            PowerBlueServerBackgroundWorker.CancelAsync()


        End If

        Dim answer As MsgBoxResult
        answer = MsgBox("Do you want to quit now?", MsgBoxStyle.YesNo)
        If answer = MsgBoxResult.Yes Then
            exitPowerPoint()
            MsgBox("Terminating program")
            End
        End If


    End Sub

    Private Sub minimizePowerBlueServerApp()
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub closeAlreadyOpenedPowerPointApp()

        pptAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")


        If Not TypeName(pptAppObj) = "Empty" Then
            pptAppObj.Quit()
        Else

        End If

    End Sub

    Private Function isPowerPointAppRunning() As Boolean

        pptAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")


        If Not TypeName(pptAppObj) = "Empty" Then

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
            MessageBox.Show(msg)
            PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Stopped.")
            StopServerButton.Enabled = False
            StartServerButton.Enabled = False
            BrowsePptButton.Enabled = True
            pptAppObj = Nothing
            presentation = Nothing
        ElseIf e.Cancelled Then
            ' Next, handle the case where the user canceled the  
            ' operation. 
            ' Note that due to a race condition in  
            ' the DoWork event handler, the Cancelled 
            ' flag may not have been set, even though 
            ' CancelAsync was called.
            ' The user canceled the operation.
            MessageBox.Show("Operation was canceled")
            PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Stopped.")
            StopServerButton.Enabled = False
            StartServerButton.Enabled = False
            BrowsePptButton.Enabled = True
            pptAppObj = Nothing
            presentation = Nothing
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
        ' MsgBox("Current Slide = " & getCurrentSlidePosition())
        'Thread.Sleep(3000)
        'controlPowerPointWithCommand("Next")
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
        Thread.Sleep(3000)
        controlPowerPointWithCommand("G001")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("G001")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("G002")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("G002")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("G003")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("G005")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Stop")
        Thread.Sleep(3000)
        controlPowerPointWithCommand("Exit")

    End Sub

   


    Private Sub startBluetoothServerInANewThread()


        PowerBlueLogTextBox.AppendText(vbCrLf & "Power Blue Server Started...")
        StartServerButton.Enabled = False
        BrowsePptButton.Enabled = False
        StopServerButton.Enabled = True


        Dim MyServiceUuid As Guid
        Dim lsnr As BluetoothListener
        'Dim received(1024) As Byte
        Dim received(3) As Byte




        MyServiceUuid = New Guid("{94f39d29-7d6d-437d-973b-fba39e49d4ee}")

        'MyServiceUuid = Guid.NewGuid()

        'MsgBox(MyServiceUuid.ToString)

        lsnr = New BluetoothListener(MyServiceUuid)
        lsnr.Start()
        'MsgBox("Server is Listening")
        'MsgBox("Server started, waiting for clients")

        blueToothClient = lsnr.AcceptBluetoothClient()
        'MsgBox("Client Connected")
        PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Client Connected" & vbCrLf & vbCrLf)

        streamRecievedFromBtClient = blueToothClient.GetStream()
        streamRecievedFromBtClient.BeginRead(received, 0, received.Length, New AsyncCallback(AddressOf ReadCallBack), received)





    End Sub


    Private Sub ReadCallBack(ar As IAsyncResult)
        'Dim received(1024) As Byte
        Dim received(3) As Byte
        Dim pptControllingCommandFull As String
        Dim pptControllingCommand As String
        pptControllingCommandFull = Nothing


        Try

            If ((ar IsNot Nothing) AndAlso ar.IsCompleted) Then

                received = ar.AsyncState
                pptControllingCommandFull = System.Text.UTF8Encoding.ASCII.GetString(received)
                'The pptControllingCommandFull is coming as string with 1024 bytes. SO i need to trim the string to get the valid string length
                If (pptControllingCommandFull IsNot Nothing) Then
                    pptControllingCommand = pptControllingCommandFull.Trim()

                    controlPowerPointWithCommand(pptControllingCommand)
                End If


                If (blueToothClient IsNot Nothing) Then

                    If (streamRecievedFromBtClient IsNot Nothing) Then

                        streamRecievedFromBtClient.Flush()
                        streamRecievedFromBtClient.BeginRead(received, 0, received.Length, New AsyncCallback(AddressOf ReadCallBack), received)


                    End If
                End If


            End If

        Catch ex As Exception
            MsgBox("Can't load Web page" & vbCrLf & ex.Message)
        End Try


    End Sub


    Private Sub controlPowerPointWithCommand(pptControllingCommand As String)
        'MsgBox("Command:" & pptControllingCommand & ":")

        If ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Open")) AndAlso (pptControllingCommand.EndsWith("Open"))) Then
            'MsgBox("Command Recieved: Open")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Open" & vbCrLf & vbCrLf)
            'Do the coding here
            openPowerPoint(pptFileName)
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Exit")) AndAlso (pptControllingCommand.EndsWith("Exit"))) Then
            'MsgBox("Command Recieved: Exit")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Exit" & vbCrLf & vbCrLf)
            exitPowerPoint()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Strt")) AndAlso (pptControllingCommand.EndsWith("Strt"))) Then
            'MsgBox("Command Recieved: Strt")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Strt" & vbCrLf & vbCrLf)
            'Do the coding here
            startPowerPointSlideShow()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Stop")) AndAlso (pptControllingCommand.EndsWith("Stop"))) Then
            'MsgBox("Command Recieved: Stop")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Stop" & vbCrLf & vbCrLf)
            'Do the coding here
            stopPowerPointSlideShow()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Rsrt")) AndAlso (pptControllingCommand.EndsWith("Rsrt"))) Then
            'MsgBox("Command Recieved: Rsrt")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Rsrt" & vbCrLf & vbCrLf)
            'Do the coding here
            restartPowerPointFromFirstSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Prev")) AndAlso (pptControllingCommand.EndsWith("Prev"))) Then
            'MsgBox("Command Recieved: Prev")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Prev" & vbCrLf & vbCrLf)
            movePowerPointToPreviousSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Next")) AndAlso (pptControllingCommand.EndsWith("Next"))) Then
            'MsgBox("Command Recieved: Next")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Next" & vbCrLf & vbCrLf)
            movePowerPointToNextSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Frst")) AndAlso (pptControllingCommand.EndsWith("Frst"))) Then
            'MsgBox("Command Recieved: Frst")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Frst" & vbCrLf & vbCrLf)
            movePowerPointToFirstSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Last")) AndAlso (pptControllingCommand.EndsWith("Last"))) Then
            'MsgBox("Command Recieved: Last")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Last" & vbCrLf & vbCrLf)
            movePowerPointToLastSlide()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Whit")) AndAlso (pptControllingCommand.EndsWith("Whit"))) Then
            'MsgBox("Command Recieved: Whit")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Whit" & vbCrLf & vbCrLf)
            'Do the coding here
            displayWhiteBackGround()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("Blac")) AndAlso (pptControllingCommand.EndsWith("Blac"))) Then
            'MsgBox("Command Recieved: Blac")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client: Blac" & vbCrLf & vbCrLf)
            'Do the coding here
            displayBlackBackGround()
        ElseIf ((pptControllingCommand.Length() = 4) AndAlso (pptControllingCommand.StartsWith("G"))) Then
            'MsgBox("Command Recieved: Blac")
            PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Command Recieved From Client:" & pptControllingCommand & vbCrLf & vbCrLf)
            'Do the coding here
            goToTheSlideNum(pptControllingCommand)
        End If


    End Sub


    Private Sub openPowerPoint(pptFileName As String)

        If (pptAppObj Is Nothing) Then
            'pptAppObj = CreateObject("powerpoint.application")
            pptAppObj = New PowerPoint.Application
            If (presentation Is Nothing) Then
                presentation = pptAppObj.Presentations.Open(FileName:=pptFileName, ReadOnly:=True)
                'pptAppObj.Visible = True
            End If

        End If

    End Sub

    Private Sub exitPowerPoint()

        If (presentation IsNot Nothing) Then
            presentation.Close()
            presentation = Nothing
        End If

        If (pptAppObj IsNot Nothing) Then
            'presentation.Close()
            pptAppObj.Quit()
            pptAppObj = Nothing
        End If



    End Sub

    Private Sub startPowerPointSlideShow()
        If (presentation IsNot Nothing) Then
            presentation.SlideShowSettings.Run()
            logCurrentSlideNumberShown()
        End If
    End Sub

    Private Sub stopPowerPointSlideShow()
        If (presentation IsNot Nothing) Then
            logCurrentSlideNumberShown()
            presentation.SlideShowWindow.View.Exit()
            'presentation.Close()
            'pptAppObj.Quit()

        End If
    End Sub

    Private Sub restartPowerPointFromFirstSlide()
        If (presentation IsNot Nothing) Then
            presentation.SlideShowWindow.View.First()
            logCurrentSlideNumberShown()
        End If
    End Sub

    Private Sub movePowerPointToPreviousSlide()
        If (presentation IsNot Nothing) Then
            If (getCurrentSlidePosition() > 1) Then
                presentation.SlideShowWindow.View.Previous()
                logCurrentSlideNumberShown()
            Else
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Cannot Move to Previous Slide. You are currently viewing the First Slide Num : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
            End If
        End If
    End Sub

    Private Sub movePowerPointToNextSlide()
        If (presentation IsNot Nothing) Then
            If (getCurrentSlidePosition() < getNumberOfSlidesInPresentation()) Then
                presentation.SlideShowWindow.View.Next()
                logCurrentSlideNumberShown()
            Else
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Cannot Move to Next Slide. You are currently viewing the Last Slide Num : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
            End If
        End If
    End Sub

    Private Sub movePowerPointToFirstSlide()
        If (presentation IsNot Nothing) Then
            If (getCurrentSlidePosition() > 1) Then
                presentation.SlideShowWindow.View.First()
                logCurrentSlideNumberShown()
            Else
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "You are currently viewing the First Slide Num Only : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
            End If

        End If

    End Sub

    Private Sub movePowerPointToLastSlide()
        If (presentation IsNot Nothing) Then
            If (getCurrentSlidePosition() < getNumberOfSlidesInPresentation()) Then
                presentation.SlideShowWindow.View.Last()
                logCurrentSlideNumberShown()
            Else
                PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "You are currently viewing the Last Slide Num Only: " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
            End If

        End If
    End Sub

    Private Sub movePowerPointToSlideNum(slideNumToGoIntLocal As Integer)
        If (presentation IsNot Nothing) Then
            presentation.SlideShowWindow.View.GotoSlide(slideNumToGoIntLocal)
            logCurrentSlideNumberShown()
        End If
    End Sub

    Private Sub displayWhiteBackGround()
        If (presentation IsNot Nothing) Then
            presentation.SlideShowWindow.View.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen
            logCurrentSlideNumberShown()
        End If

    End Sub

    Private Sub displayBlackBackGround()
        If (presentation IsNot Nothing) Then
            presentation.SlideShowWindow.View.State = PowerPoint.PpSlideShowState.ppSlideShowBlackScreen
            logCurrentSlideNumberShown()
        End If
    End Sub

    Private Sub goToTheSlideNum(pptControllingCommandLocal As String)
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


    End Sub

    Private Function getNumberOfSlidesInPresentation() As Integer
        If (presentation IsNot Nothing) Then
            Return presentation.Slides.Count()
        Else
            Return 0
        End If
    End Function

    Private Function getCurrentSlidePosition() As Integer
        If (presentation IsNot Nothing) Then
            Return presentation.SlideShowWindow.View.CurrentShowPosition
        Else
            Return 0
        End If
    End Function

    Private Sub logCurrentSlideNumberShown()
        PowerBlueLogTextBox.AppendText(vbCrLf & vbCrLf & "Current Slide No Showing : " & getCurrentSlidePosition() & vbCrLf & vbCrLf)
    End Sub


End Class