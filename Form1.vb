Imports System
Imports System.IO.Ports
Imports System.Timers
Public Class Form1
    Dim TxDATA(26) As Char
    Dim ReadFlag As Boolean = False             '接收Flag 
    Dim Buffer As String = ""
    Dim Input As String = ""


    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If SerialPort.IsOpen = False Then Call Enable(False)

        TxDATA(0) = "<"                         'Start Bits (Iminate UART 's Start Bits 微處理機第二章會教)
        TxDATA(1) = "A"
        TxDATA(2) = ">"
        For i = 3 To 26                         'Initialize
            TxDATA(i) = "0"
        Next
    End Sub
    Private Sub ComboBox_DropDown(sender As Object, e As EventArgs) Handles ComboBox.DropDown
        ComboBox.Items.Clear()
        For Each ports As String In My.Computer.Ports.SerialPortNames
            ComboBox.Items.Add(ports)
        Next
        ComboBox.Sorted = True                                           '讓Combobox照順序排 : True
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick, Timer1.Tick
        Call CurrentTime()
        Call RecieveDATA()
    End Sub
    Private Sub CurrentTime()
        Dim tHour, tMinute, tSecond As Integer
        Dim Current As String
        Current = "Current Time:"

        tHour = Hour(Now())
        tMinute = Minute(Now()) + 10
        tSecond = Second(Now())

        If tMinute >= 60 Then
            tMinute -= 60
            tHour += 1
        End If

        If tHour < 10 Then
            Current &= "0" & tHour & ":"
        Else
            Current &= tHour & ":"
        End If

        If tMinute < 10 Then
            Current &= "0" & tMinute & ":"
        Else
            Current &= tMinute & ":"
        End If

        If tSecond < 10 Then
            Current &= "0" & tSecond
        Else
            Current &= tSecond
        End If

        CurrentTimeBox.Text = Current
    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If SerialPort.IsOpen = True Then SerialPort.Close()
        End
    End Sub
    Private Sub OpenButton_Click(sender As Object, e As EventArgs) Handles OpenButton.Click
        If SerialPort.IsOpen = False Then
            SerialPort.Open()
            DeviceStatus.BackColor = Color.FromArgb(0, 255, 0)          ' VB.net 的 RGB超麻煩 哭啊 = =
            Call Enable(True)
        End If
    End Sub
    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        If SerialPort.IsOpen = True Then
            SerialPort.Close()
            DeviceStatus.BackColor = Color.FromArgb(255, 0, 0)
            Call Enable(False)
        End If
    End Sub
    Private Sub Enable(enableFlag As Boolean)
        WriteTextBox.Enabled = enableFlag
        ReadTextBox.Enabled = enableFlag
        WriteButton.Enabled = enableFlag
        ReadButton.Enabled = enableFlag
    End Sub
    Private Sub WriteButton_Click(sender As Object, e As EventArgs) Handles WriteButton.Click
        If SerialPort.IsOpen = True Then
            If Len(WriteTextBox.Text) = 4 _
                    And ((Mid(WriteTextBox.Text, 1, 1) >= "A" And Mid(WriteTextBox.Text, 1, 1) <= "F") _
                    And (Mid(WriteTextBox.Text, 2, 1) >= "A" And Mid(WriteTextBox.Text, 2, 1) <= "F") _
                    And (Mid(WriteTextBox.Text, 3, 1) >= "A" And Mid(WriteTextBox.Text, 3, 1) <= "F") _
                    And (Mid(WriteTextBox.Text, 4, 1) >= "A" And Mid(WriteTextBox.Text, 4, 1) <= "F")) Then

                TxDATA(7) = "1"                                          'Function ---  1 : Write   2 : Read
                TxDATA(8) = Mid(WriteTextBox.Text, 1, 1)
                TxDATA(9) = Mid(WriteTextBox.Text, 2, 1)
                TxDATA(10) = Mid(WriteTextBox.Text, 3, 1)
                TxDATA(11) = Mid(WriteTextBox.Text, 4, 1)
                SerialPort.Write(TxDATA)

                ReadFlag = False
            Else
                MsgBox("Not 4 Bits Hex Format", vbOKOnly, "110")     'MsgBox "content" ,  "button" ,  "title"
            End If
        End If
    End Sub

    Private Sub ReadTextBox_TextChanged(sender As Object, e As EventArgs) Handles ReadTextBox.TextChanged
        If SerialPort.IsOpen = True Then
            TxDATA(7) = "2"                                             'Function ---  1 : Write   2 : Read
            SerialPort.Write(TxDATA)
            ReadFlag = True
        End If
    End Sub
    Private Sub RecieveDATA()
        If SerialPort.IsOpen = True And ReadFlag = True Then
            Do
                Buffer = SerialPort.ReadLine()
                If Buffer Is Nothing Then
                    Exit Do
                Else
                    Input &= Buffer & vbCrLf
                End If
            Loop
        End If
    End Sub
End Class