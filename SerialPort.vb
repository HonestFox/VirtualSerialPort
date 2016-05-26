Imports System.IO.Ports
Imports System.Threading
Imports System.Security.Permissions

'允许在程序内部各个线程对于窗体控件进行操作
<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class Form1

    Dim thread1, thread2, thread3 As Threading.Thread
    Dim ThreadCount As Integer = 0
    Dim DTR_TIME, DTR_DUTYRATIO As Integer
    Dim RTS_TIME, RTS_DUTYRATIO As Integer


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SerialPort1.Encoding = System.Text.Encoding.Default
        ComboBox1.Items.Clear()   '先清空
        For Each port As String In My.Computer.Ports.SerialPortNames  '把所有串口号遍历一遍 把所有串口号放入ComboBox1中
            ComboBox1.Items.Add(port)
        Next

        ComboBox1.SelectedIndex = 0  '默认为第一个搜索到的串口
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '打开串口

        With SerialPort1
            .BaudRate = 9600
            .StopBits = 1
            .DataBits = 8
            .Parity = IO.Ports.Parity.None
            .PortName = ComboBox1.SelectedItem.ToString
        End With

        SerialPort1.Open()
        MsgBox("串口打开")
        Button1.Enabled = False    '避免反复打开同一个串口
        Button2.Enabled = True
        Button3.Enabled = True

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '关闭串口
        SerialPort1.Close()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '发送数据
        If TextBox1.Text = "" Then
            TextBox1.Text = " "
        End If

        If TextBox5.Text = "" Then
            TextBox5.Text = 1000
        End If
        If (CheckBox5.Checked = True) Then   '自动发送
            thread3 = New Threading.Thread(AddressOf _AutoSendMSG)
            thread3.Start()
            ThreadCount += 1
        Else
            SerialPort1.Write(Now + " " + SerialPort1.PortName + vbNewLine)  '发送
            SerialPort1.Write(TextBox1.Text)  '发送
            TextBox1.Clear()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        '允许线程对窗口控件进行操作
        Control.CheckForIllegalCrossThreadCalls = False

        Dim str As String
        str = SerialPort1.ReadExisting
        TextBox2.Text += str
        TextBox2.Text += vbNewLine

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        'DTR
        If CheckBox1.Checked = True Then
            SerialPort1.DtrEnable = True
            Button4.BackColor = Color.White
        Else
            SerialPort1.DtrEnable = False
            Button4.BackColor = Color.Black

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        'RTS
        If CheckBox2.Checked = True Then
            SerialPort1.RtsEnable = True
            Button5.BackColor = Color.White
        Else
            SerialPort1.RtsEnable = False
            Button5.BackColor = Color.Black
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'DTR指示灯_1   
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        'DTR指示灯_2
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'RTS指示灯_1
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        'RTS指示灯_2
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        '产生DTR波形按钮
        If TextBox3.Text = "" Then
            TextBox3.Text = 3  '默认周期
        End If

        If TextBox4.Text = "" Then
            TextBox4.Text = 50
        End If

        DTR_TIME = TextBox3.Text
        DTR_DUTYRATIO = TextBox4.Text

        thread1 = New Threading.Thread(AddressOf DTRWaveGenerate)
        thread1.Start()
        ThreadCount += 1
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        '产生RTS波形按钮

        If TextBox3.Text = "" Then
            TextBox3.Text = 3
        End If

        If TextBox4.Text = "" Then
            TextBox4.Text = 50
        End If
        RTS_TIME = TextBox3.Text
        RTS_DUTYRATIO = TextBox4.Text
        thread2 = New Threading.Thread(AddressOf RTSWaveGenerate)
        thread2.Start()
        ThreadCount += 1
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        '释放DTR线程按钮

        thread1.Abort()
        ThreadCount -= 1

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        '释放RTS线程按钮
        thread2.Abort()
        ThreadCount -= 1
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        '周期
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        '占空比
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub MDIForm_Unload(Cancel As Integer)
        thread1.Abort()
        thread2.Abort()
        End
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        '产生DTR方波
        If CheckBox3.Checked = True Then

            If TextBox3.Text = "" Then
                TextBox3.Text = 3  '默认周期
            End If

            If TextBox4.Text = "" Then
                TextBox4.Text = 50
            End If

            DTR_TIME = TextBox3.Text
            DTR_DUTYRATIO = TextBox4.Text

            thread1 = New Threading.Thread(AddressOf DTRWaveGenerate)
            thread1.Start()
            ThreadCount += 1

        Else
            Button4.BackColor = Color.Black
            thread1.Abort()
            ThreadCount -= 1
        End If

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        '产生RTS方波
        If CheckBox4.Checked = True Then
            If TextBox3.Text = "" Then
                TextBox3.Text = 3
            End If

            If TextBox4.Text = "" Then
                TextBox4.Text = 50
            End If
            RTS_TIME = TextBox3.Text
            RTS_DUTYRATIO = TextBox4.Text
            thread2 = New Threading.Thread(AddressOf RTSWaveGenerate)
            thread2.Start()
            ThreadCount += 1
        Else
            Button5.BackColor = Color.Black
            thread2.Abort()
            ThreadCount -= 1
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        '自动发送按钮
        'If CheckBox5.Checked = True Then
        '    If (TextBox5.Text = "") Then
        '        TextBox5.Text = "1000"
        '    End If
        '    If (CheckBox5.Checked = True) Then   '自动发送
        '        thread3 = New Threading.Thread(AddressOf _AutoSendMSG)
        '        thread3.Start()
        '        ThreadCount += 1
        '    Else
        '        SerialPort1.Write(Now + " " + SerialPort1.PortName + vbNewLine)  '发送
        '        SerialPort1.Write(TextBox1.Text)  '发送
        '        TextBox1.Clear()
        '    End If
        'Else
        If (TextBox5.Text = "") Then
            '       TextBox5.Text = "1000"
        End If
        If CheckBox5.Checked = False Then
            If ThreadCount > 0 Then
                thread3.Abort()
                ThreadCount -= 1
            End If
        End If

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Sub DTRWaveGenerate()
        DTR_TIME *= 1000
        Dim HighLevelTime = DTR_TIME * DTR_DUTYRATIO / 100
        Dim LowLevelTime = DTR_TIME - HighLevelTime

        While True
            SerialPort1.DtrEnable = True
            Button4.BackColor = Color.White
            Thread.Sleep(HighLevelTime)

            SerialPort1.DtrEnable = False
            Button4.BackColor = Color.Black
            Thread.Sleep(LowLevelTime)
        End While
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        '关闭时调用这个

        If ThreadCount <> 0 Then
            If MsgBox("正在输出信号，是否强制停止并退出？! " & vbNewLine, vbExclamation + vbYesNo, "强制退出？") = vbYes Then
                MsgBox("已强制终止！")
                Try
                    thread1.Abort()
                    thread2.Abort()
                Catch ex As Exception
                End Try
            Else
                e.Cancel = 1
            End If
    End If

    End Sub

    Sub RTSWaveGenerate()
        RTS_TIME *= 1000
        Dim HighLevelTime = RTS_TIME * RTS_DUTYRATIO / 100
        Dim LowLevelTime = RTS_TIME - HighLevelTime

        While True
            SerialPort1.RtsEnab   le = True
            Button5.BackColor = Color.White
            Thread.Sleep(HighLevelTime)

            SerialPort1.RtsEnable = False
            Button5.BackColor = Color.Black
            Thread.Sleep(LowLevelTime)
        End While
    End Sub

    Sub _AutoSendMSG()
        Dim Interval As Integer
        Interval = TextBox5.Text
        While True
            Try
                SerialPort1.Write(Now + " " + SerialPort1.PortName + vbNewLine)  '发送
            Catch
            End Tryg
            If TextBox1.Text = "" Then
                TextBox1.Text = " "
            End If

            SerialPort1.Write(TextBox1.Text)  '发送
            ' TextBox1.Clear()
            Thread.Sleep(Interval)
        End While

    End Sub

    Private Sub SerialPort1_PinChanged(sender As Object, e As SerialPinChangedEventArgs) Handles SerialPort1.PinChanged
        '扩展
        '   If SerialPort1.
    End Sub
End Class
