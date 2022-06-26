'----Function----
'TxDATA(7)  ->  7seg = 1 / 4bits_7seg = 2 / LED = 3
'TxDATA(8), TxDATA(9)  -> 7seg_val
'TxDATA(8), TxDATA(9), TxDATA(10), TxDATA(11)  -> 4bits_7seg_val

'TxDATA(12), TxDATA(13), TxDATA(14), TxDATA(15),
'TxDATA(16), TxDATA(17), TxDATA(18), TxDATA(19)    ->  LED
Public Class Form1
    Dim TxDATA(26) As Char

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If SerialPort.IsOpen = False Then Call Enable(False)

        TxDATA(0) = "<"                         'Start Bits (Iminate UART 's Start Bits 微處理機第二章會教)
        TxDATA(1) = "A"
        TxDATA(2) = ">"
        For i = 3 To 26                         'Initialize
            TxDATA(i) = "0"
        Next
    End Sub
    Private Sub Enable(enableFlag As Boolean)
        WriteTextBox.Enabled = enableFlag
        WriteButton.Enabled = enableFlag
        AllLightsON.Enabled = enableFlag
        AllLightsOFF.Enabled = enableFlag
        LED0.Enabled = enableFlag
        LED1.Enabled = enableFlag
        LED2.Enabled = enableFlag
        LED3.Enabled = enableFlag
        LED4.Enabled = enableFlag
        LED5.Enabled = enableFlag
        LED6.Enabled = enableFlag
        LED7.Enabled = enableFlag
    End Sub
    Private Sub ComboBox_DropDown(sender As Object, e As EventArgs) Handles ComboBox.DropDown
        ComboBox.Items.Clear()
        For Each ports As String In My.Computer.Ports.SerialPortNames
            ComboBox.Items.Add(ports)
        Next
        ComboBox.Sorted = True                                           '讓Combobox照順序排 : True
    End Sub
    Private Sub WriteButton_Click(sender As Object, e As EventArgs) Handles WriteButton.Click
        If SerialPort.IsOpen = True Then
            If Len(WriteTextBox.Text) = 4 _
                    And ((Mid(WriteTextBox.Text, 1, 1) >= "0" And Mid(WriteTextBox.Text, 1, 1) <= "9") _
                    And (Mid(WriteTextBox.Text, 2, 1) >= "0" And Mid(WriteTextBox.Text, 2, 1) <= "9") _
                    And (Mid(WriteTextBox.Text, 3, 1) >= "0" And Mid(WriteTextBox.Text, 3, 1) <= "9") _
                    And (Mid(WriteTextBox.Text, 4, 1) >= "0" And Mid(WriteTextBox.Text, 4, 1) <= "9")) Then

                TxDATA(7) = "1"                                          'Function ---  1 : Write   2 : Read
                TxDATA(8) = Mid(WriteTextBox.Text, 1, 1)
                TxDATA(9) = Mid(WriteTextBox.Text, 2, 1)
                TxDATA(10) = Mid(WriteTextBox.Text, 3, 1)
                TxDATA(11) = Mid(WriteTextBox.Text, 4, 1)
                SerialPort.Write(TxDATA)
            Else
                WriteTextBox.Text = ""
            End If
        End If
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
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If SerialPort.IsOpen = True Then SerialPort.Close()
        End
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Call CurrentTime()
        Call LED_Checked()
    End Sub
    Private Sub CurrentTime()
        Dim tHour, tMinute, tSecond As Integer
        Dim Current As String
        Current = "Current Time:"

        tHour = Hour(Now())
        tMinute = Minute(Now())
        tSecond = Second(Now())

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
    Private Sub LED_Checked()
        For i = 12 To 19
            TxDATA(i) = "0"
        Next
        TxDATA(7) = "3"
        If LED7.Checked = True Then TxDATA(12) = "1"
        If LED6.Checked = True Then TxDATA(13) = "1"
        If LED5.Checked = True Then TxDATA(14) = "1"
        If LED4.Checked = True Then TxDATA(15) = "1"
        If LED3.Checked = True Then TxDATA(16) = "1"
        If LED2.Checked = True Then TxDATA(17) = "1"
        If LED1.Checked = True Then TxDATA(18) = "1"
        If LED0.Checked = True Then TxDATA(19) = "1"
        SerialPort.Write(TxDATA)
    End Sub
    Private Sub AllLightsON_Click(sender As Object, e As EventArgs) Handles AllLightsON.Click
        If SerialPort.IsOpen = True Then
            Call LED_State(True)
            Dim f As String = "1"
            TxDATA(7) = "3"
            TxDATA(6) = f
            TxDATA(12) = f
            TxDATA(13) = f
            TxDATA(14) = f
            TxDATA(15) = f
            TxDATA(16) = f
            TxDATA(17) = f
            TxDATA(18) = f
            TxDATA(19) = f
            SerialPort.Write(TxDATA)
        End If
    End Sub
    Private Sub AllLightsOFF_Click(sender As Object, e As EventArgs) Handles AllLightsOFF.Click
        If SerialPort.IsOpen = True Then
            Call LED_State(False)
            Dim f As String = "0"
            TxDATA(7) = "3"
            TxDATA(12) = f
            TxDATA(13) = f
            TxDATA(14) = f
            TxDATA(15) = f
            TxDATA(16) = f
            TxDATA(17) = f
            TxDATA(18) = f
            TxDATA(19) = f
            SerialPort.Write(TxDATA)
        End If
    End Sub
    Private Sub LED_State(StateFlag As Boolean)
        LED0.Checked = StateFlag
        LED1.Checked = StateFlag
        LED2.Checked = StateFlag
        LED3.Checked = StateFlag
        LED4.Checked = StateFlag
        LED5.Checked = StateFlag
        LED6.Checked = StateFlag
        LED7.Checked = StateFlag
    End Sub
End Class
