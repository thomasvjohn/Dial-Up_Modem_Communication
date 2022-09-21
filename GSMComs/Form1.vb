Imports System.IO
Public Class Form1

#Region " GSM Modem Site "

    Private Sub OpenCOM1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            With SerialPort1
                If .IsOpen = False Then
                    .BaudRate = Val(ComboBox1.SelectedItem)
                    .DataBits = Val(TextBox1.Text)
                    If ComboBox2.SelectedItem.Contains("True") Then
                        .DiscardNull = True
                    Else
                        .DiscardNull = False
                    End If
                    If ComboBox3.SelectedItem.Contains("True") Then
                        .DtrEnable = True
                    Else
                        .DtrEnable = False
                    End If
                    If ComboBox4.SelectedItem.Contains("RequestToSendXOnXOff") Then
                        .Handshake = Ports.Handshake.RequestToSendXOnXOff
                    ElseIf ComboBox4.SelectedItem.Contains("XOnXOff") Then
                        .Handshake = Ports.Handshake.XOnXOff
                    ElseIf ComboBox4.SelectedItem.Contains("RequestToSend") Then
                        .Handshake = Ports.Handshake.RequestToSend
                    ElseIf ComboBox4.SelectedItem.Contains("None") Then
                        .Handshake = Ports.Handshake.None
                    End If
                    .ReadBufferSize = Val(TextBox3.Text)
                    If ComboBox6.SelectedItem.Contains("True") Then
                        .RtsEnable = True
                    Else
                        .RtsEnable = False
                    End If
                    .WriteBufferSize = Val(TextBox4.Text)

                    Application.DoEvents()

                    .Open()
                    Application.DoEvents()

                    Timer1.Enabled = True
                    Button2.Enabled = False
                    Button3.Enabled = True
                    Button4.Enabled = True
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Dim MR1 As String = Nothing
    Dim HangStep1 As Integer = 0
    Private Sub COM1ReadTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        With SerialPort1
            If .BytesToRead > 0 Then
                MR1 = .ReadLine
                ListBox1.Items.Insert(0, Now.ToString & " Received: " & MR1)
                If MR1.Contains("NO CARRIER") = True Then
                    HangStep1 = 4
                    HangTimer1.Enabled = True
                ElseIf MR1.Contains("NO DIALTONE") = True Then
                    HangStep1 = 4
                    HangTimer1.Enabled = True
                ElseIf MR1.Contains("BUSY") = True Then
                    HangStep1 = 4
                    HangTimer1.Enabled = True
                ElseIf MR1.Contains("ERROR") = True Then
                    HangStep1 = 4
                    HangTimer1.Enabled = True
                ElseIf MR1.Contains("CONNECT") = True Then
                    System.Threading.Thread.Sleep(10)
                End If
            End If
        End With

    End Sub

    Private Sub HangTimer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HangTimer1.Tick
        With SerialPort1
            HangTimer1.Enabled = False
            Application.DoEvents()
            Select Case HangStep1
                Case 1
                    HangStep1 = 2
                    HangTimer1.Enabled = True
                    Application.DoEvents()
                Case 2
                    'Updating Label and Clearing Buffer
                    Application.DoEvents()
                    Do While (.BytesToWrite > 0)
                    Loop
                    .WriteLine("+++" & vbCr)
                    HangStep1 = 3
                    HangTimer1.Enabled = True
                    Application.DoEvents()
                Case 3
                    'Enabling DTR
                    .DtrEnable = True
                    HangStep1 = 4
                    HangTimer1.Enabled = True
                    Application.DoEvents()
                Case 4
                    'Disabling DTR
                    .DtrEnable = False
                    HangStep1 = 5
                    HangTimer1.Enabled = True
                    Application.DoEvents()
                Case 5
                    'Clearing Buffer
                    Do While (.BytesToWrite > 0)
                    Loop
                    'Hanging the Call
                    .WriteLine("ATH" & vbCr)
                    HangStep1 = 6
                    HangTimer1.Enabled = True
                    Application.DoEvents()
                Case 6
                    'Updating Status
                    Application.DoEvents()
                    Do While (.BytesToWrite > 0)
                    Loop
                    .WriteLine("ATS0=1" & vbCr)
                    'ModemCOM.RtsEnable = False
                    HangStep1 = 0
                    Application.DoEvents()
            End Select
        End With
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If TextBox2.Text.Contains("+++") Then
            SerialPort1.WriteLine(TextBox2.Text.Trim)
        Else
            SerialPort1.WriteLine(TextBox2.Text & vbCr)
        End If

        ListBox1.Items.Insert(0, "Send:" & TextBox2.Text)
        TextBox2.Text = ""
        Application.DoEvents()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        With SerialPort1
            If .IsOpen = True Then
                Timer1.Enabled = False
                .Close()
                Application.DoEvents()
                Button2.Enabled = True
                Button3.Enabled = False
                Button4.Enabled = False
            End If
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        With SerialPort1
            If ComboBox3.SelectedItem.Contains("True") AndAlso .DtrEnable = False Then
                .DtrEnable = True
            ElseIf ComboBox3.SelectedItem.Contains("False") AndAlso .DtrEnable = True Then
                .DtrEnable = False
            End If
        End With
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim SR As New StreamReader(OpenFileDialog1.FileName)
                Do Until SR.Peek = -1
                    Dim DT As String = SR.ReadLine
                    SerialPort1.WriteLine(DT & vbCr)
                    ListBox1.Items.Insert(0, "Send:" & DT)
                    Application.DoEvents()
                Loop
                SR.Close()
            End If
        Catch ex As Exception
            ListBox1.Items.Insert(0, "ERROR:" & ex.Message)
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Application.DoEvents()
        Do While (SerialPort1.BytesToWrite > 0)
        Loop
    End Sub

#End Region

#Region " USB Modem Master "

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            With SerialPort2
                If .IsOpen = False Then
                    .BaudRate = Val(ComboBox10.SelectedItem)
                    .DataBits = Val(TextBox8.Text)
                    If ComboBox9.SelectedItem.Contains("True") Then
                        .DiscardNull = True
                    Else
                        .DiscardNull = False
                    End If
                    If ComboBox8.SelectedItem.Contains("True") Then
                        .DtrEnable = True
                    Else
                        .DtrEnable = False
                    End If
                    If ComboBox7.SelectedItem.Contains("RequestToSendXOnXOff") Then
                        .Handshake = Ports.Handshake.RequestToSendXOnXOff
                    ElseIf ComboBox7.SelectedItem.Contains("XOnXOff") Then
                        .Handshake = Ports.Handshake.XOnXOff
                    ElseIf ComboBox7.SelectedItem.Contains("RequestToSend") Then
                        .Handshake = Ports.Handshake.RequestToSend
                    ElseIf ComboBox7.SelectedItem.Contains("None") Then
                        .Handshake = Ports.Handshake.None
                    End If
                    .ReadBufferSize = Val(TextBox5.Text)
                    If ComboBox5.SelectedItem.Contains("True") Then
                        .RtsEnable = True
                    Else
                        .RtsEnable = False
                    End If
                    .WriteBufferSize = Val(TextBox5.Text)
                    Application.DoEvents()

                    .Open()
                    Application.DoEvents()

                    Timer2.Enabled = True
                    Button7.Enabled = False
                    Button6.Enabled = True
                    Button5.Enabled = True
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Dim MR2 As String = Nothing
    Dim HangStep2 As Integer = 0
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        With SerialPort2
            If .BytesToRead > 0 Then
                MR2 = .ReadLine
                ListBox2.Items.Insert(0, Now.ToString & " Received: " & MR2)

                If MR2.Contains("NO CARRIER") = True Then
                    HangStep2 = 4
                    HangTimer2.Enabled = True
                ElseIf MR2.Contains("NO DIALTONE") = True Then
                    HangStep2 = 4
                    HangTimer2.Enabled = True
                ElseIf MR2.Contains("BUSY") = True Then
                    HangStep2 = 4
                    HangTimer2.Enabled = True
                ElseIf MR2.Contains("ERROR") = True Then
                    HangStep2 = 4
                    HangTimer2.Enabled = True
                ElseIf MR2.Contains("CONNECT") = True Then
                    System.Threading.Thread.Sleep(10)
                End If
            End If
        End With

    End Sub

    Private Sub HangTimer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HangTimer2.Tick
        With SerialPort2
            HangTimer2.Enabled = False
            Application.DoEvents()
            Select Case HangStep2
                Case 1
                    HangStep2 = 2
                    HangTimer2.Enabled = True
                    Application.DoEvents()
                Case 2
                    'Updating Label and Clearing Buffer
                    Application.DoEvents()
                    Do While (.BytesToWrite > 0)
                    Loop
                    .WriteLine("+++" & vbCr)
                    HangStep2 = 3
                    HangTimer2.Enabled = True
                    Application.DoEvents()
                Case 3
                    'Enabling DTR
                    .DtrEnable = True
                    HangStep2 = 4
                    HangTimer2.Enabled = True
                    Application.DoEvents()
                Case 4
                    'Disabling DTR
                    .DtrEnable = False
                    HangStep2 = 5
                    HangTimer2.Enabled = True
                    Application.DoEvents()
                Case 5
                    'Clearing Buffer
                    Do While (.BytesToWrite > 0)
                    Loop
                    'Hanging the Call
                    .WriteLine("ATH" & vbCr)
                    HangStep2 = 6
                    HangTimer2.Enabled = True
                    Application.DoEvents()
                Case 6
                    'Updating Status
                    Application.DoEvents()
                    Do While (.BytesToWrite > 0)
                    Loop
                    .WriteLine("ATS0=1" & vbCr)
                    'ModemCOM.RtsEnable = False
                    HangStep2 = 0
                    Application.DoEvents()
            End Select
        End With
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox7.Text.Contains("+++") Then
            SerialPort2.WriteLine(TextBox7.Text.Trim)
        Else
            SerialPort2.WriteLine(TextBox7.Text & vbCr)
        End If

        ListBox2.Items.Insert(0, "Send:" & TextBox7.Text)
        TextBox7.Text = ""
        Application.DoEvents()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        With SerialPort2
            If .IsOpen = True Then
                Timer2.Enabled = False
                .Close()
                Application.DoEvents()
                Button7.Enabled = True
                Button6.Enabled = False
                Button5.Enabled = False
            End If
        End With
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ListBox2.Items.Clear()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        With SerialPort2
            If ComboBox8.SelectedItem.Contains("True") AndAlso .DtrEnable = False Then
                .DtrEnable = True
            ElseIf ComboBox8.SelectedItem.Contains("False") AndAlso .DtrEnable = True Then
                .DtrEnable = False
            End If
            If ComboBox5.SelectedItem.Contains("True") AndAlso .RtsEnable = False Then
                .RtsEnable = True
            ElseIf ComboBox5.SelectedItem.Contains("False") AndAlso .RtsEnable = True Then
                .RtsEnable = False
            End If
        End With
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Try
            If OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim SR As New StreamReader(OpenFileDialog2.FileName)
                Do Until SR.Peek = -1
                    Dim DT As String = SR.ReadLine
                    SerialPort2.WriteLine(DT & vbCr)
                    ListBox2.Items.Insert(0, "Send:" & DT)
                    Application.DoEvents()
                Loop
                SR.Close()
            End If
        Catch ex As Exception
            ListBox2.Items.Insert(0, "ERROR:" & ex.Message)
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Application.DoEvents()
        Do While (SerialPort2.BytesToWrite > 0)
        Loop
    End Sub
#End Region

#Region " USB Modem Site "
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Try
            With SerialPort3
                If .IsOpen = False Then
                    .BaudRate = Val(ComboBox15.SelectedItem)
                    .DataBits = Val(TextBox12.Text)
                    If ComboBox14.SelectedItem.Contains("True") Then
                        .DiscardNull = True
                    Else
                        .DiscardNull = False
                    End If
                    If ComboBox13.SelectedItem.Contains("True") Then
                        .DtrEnable = True
                    Else
                        .DtrEnable = False
                    End If
                    If ComboBox12.SelectedItem.Contains("RequestToSendXOnXOff") Then
                        .Handshake = Ports.Handshake.RequestToSendXOnXOff
                    ElseIf ComboBox12.SelectedItem.Contains("XOnXOff") Then
                        .Handshake = Ports.Handshake.XOnXOff
                    ElseIf ComboBox12.SelectedItem.Contains("RequestToSend") Then
                        .Handshake = Ports.Handshake.RequestToSend
                    ElseIf ComboBox12.SelectedItem.Contains("None") Then
                        .Handshake = Ports.Handshake.None
                    End If
                    .ReadBufferSize = Val(TextBox10.Text)
                    If ComboBox11.SelectedItem.Contains("True") Then
                        .RtsEnable = True
                    Else
                        .RtsEnable = False
                    End If
                    .WriteBufferSize = Val(TextBox9.Text)
                    Application.DoEvents()

                    .Open()
                    Application.DoEvents()

                    Timer3.Enabled = True
                    Button21.Enabled = False
                    Button20.Enabled = True
                    Button19.Enabled = True
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        With SerialPort3
            If .IsOpen = True Then
                Timer3.Enabled = False
                .Close()
                Application.DoEvents()
                Button21.Enabled = True
                Button20.Enabled = False
                Button19.Enabled = False
            End If
        End With
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Application.DoEvents()
        Do While (SerialPort3.BytesToWrite > 0)
        Loop
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Try
            If OpenFileDialog3.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim SR As New StreamReader(OpenFileDialog3.FileName)
                Do Until SR.Peek = -1
                    Dim DT As String = SR.ReadLine
                    SerialPort3.WriteLine(DT & vbCr)
                    ListBox3.Items.Insert(0, "Send:" & DT)
                    Application.DoEvents()
                Loop
                SR.Close()
            End If
        Catch ex As Exception
            ListBox3.Items.Insert(0, "ERROR:" & ex.Message)
        End Try
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        With SerialPort3
            If ComboBox13.SelectedItem.Contains("True") AndAlso .DtrEnable = False Then
                .DtrEnable = True
            ElseIf ComboBox13.SelectedItem.Contains("False") AndAlso .DtrEnable = True Then
                .DtrEnable = False
            End If
        End With
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        ListBox3.Items.Clear()
    End Sub

    Dim MR3 As String = Nothing
    Dim HangStep3 As Integer = 0
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        With SerialPort3
            If .BytesToRead > 0 Then
                MR3 = .ReadLine
                ListBox3.Items.Insert(0, Now.ToString & " Received: " & MR3)

                If MR3.Contains("NO CARRIER") = True Then
                    HangStep3 = 4
                    HangTimer3.Enabled = True
                ElseIf MR3.Contains("NO DIALTONE") = True Then
                    HangStep3 = 4
                    HangTimer3.Enabled = True
                ElseIf MR3.Contains("BUSY") = True Then
                    HangStep3 = 4
                    HangTimer3.Enabled = True
                ElseIf MR3.Contains("ERROR") = True Then
                    HangStep3 = 4
                    HangTimer3.Enabled = True
                ElseIf MR3.Contains("CONNECT") = True Then
                    System.Threading.Thread.Sleep(10)
                End If
            End If
        End With

    End Sub

    Private Sub HangTimer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HangTimer3.Tick
        With SerialPort3
            HangTimer3.Enabled = False
            Application.DoEvents()
            Select Case HangStep3
                Case 1
                    HangStep3 = 2
                    HangTimer3.Enabled = True
                    Application.DoEvents()
                Case 2
                    'Updating Label and Clearing Buffer
                    Application.DoEvents()
                    Do While (.BytesToWrite > 0)
                    Loop
                    .WriteLine("+++" & vbCr)
                    HangStep3 = 3
                    HangTimer3.Enabled = True
                    Application.DoEvents()
                Case 3
                    'Enabling DTR
                    .DtrEnable = True
                    HangStep3 = 4
                    HangTimer3.Enabled = True
                    Application.DoEvents()
                Case 4
                    'Disabling DTR
                    .DtrEnable = False
                    HangStep3 = 5
                    HangTimer3.Enabled = True
                    Application.DoEvents()
                Case 5
                    'Clearing Buffer
                    Do While (.BytesToWrite > 0)
                    Loop
                    'Hanging the Call
                    .WriteLine("ATH" & vbCr)
                    HangStep3 = 6
                    HangTimer3.Enabled = True
                    Application.DoEvents()
                Case 6
                    'Updating Status
                    Application.DoEvents()
                    Do While (.BytesToWrite > 0)
                    Loop
                    .WriteLine("ATS0=1" & vbCr)
                    'ModemCOM.RtsEnable = False
                    HangStep3 = 0
                    Application.DoEvents()
            End Select
        End With
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        SerialPort3.WriteLine(TextBox11.Text & vbCr)
        ListBox3.Items.Insert(0, "Send:" & TextBox11.Text)
        TextBox11.Text = ""
        Application.DoEvents()
    End Sub
#End Region

   
#Region " System "
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button5.Enabled = False
        Button20.Enabled = False
        Button19.Enabled = False
    End Sub
    Private Sub Form1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox2.Text <> Nothing Then
                SerialPort1.WriteLine(TextBox2.Text & vbCr)
                ListBox1.Items.Insert(0, "Send:" & TextBox2.Text)
                TextBox2.Text = ""
                Application.DoEvents()
            ElseIf TextBox7.Text <> Nothing Then
                SerialPort2.WriteLine(TextBox7.Text & vbCr)
                ListBox2.Items.Insert(0, "Send:" & TextBox7.Text)
                TextBox7.Text = ""
                Application.DoEvents()
            ElseIf TextBox11.Text <> Nothing Then
                SerialPort3.WriteLine(TextBox11.Text & vbCr)
                ListBox3.Items.Insert(0, "Send:" & TextBox11.Text)
                TextBox11.Text = ""
                Application.DoEvents()
            End If
        End If
    End Sub
    
#End Region

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        With SerialPort3
            If ComboBox11.SelectedItem.Contains("True") AndAlso .RtsEnable = False Then
                .RtsEnable = True
            ElseIf ComboBox11.SelectedItem.Contains("False") AndAlso .RtsEnable = True Then
                .RtsEnable = False
            End If
        End With
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        With SerialPort1
            If ComboBox6.SelectedItem.Contains("True") AndAlso .RtsEnable = False Then
                .RtsEnable = True
            ElseIf ComboBox6.SelectedItem.Contains("False") AndAlso .RtsEnable = True Then
                .RtsEnable = False
            End If
        End With
    End Sub
End Class
