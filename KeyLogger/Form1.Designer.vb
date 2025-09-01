<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Menu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.StartServiceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.StatusToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitServiceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu.SuspendLayout()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 1)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(354, 10)
        Me.ProgressBar1.TabIndex = 1
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenuStrip = Me.Menu
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Key Logger"
        Me.NotifyIcon1.Visible = True
        '
        'Menu
        '
        Me.Menu.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.Menu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StartServiceToolStripMenuItem, Me.ToolStripMenuItem1, Me.StatusToolStripMenuItem, Me.ToolStripSeparator1, Me.ExitServiceToolStripMenuItem})
        Me.Menu.Name = "Menu"
        Me.Menu.Size = New System.Drawing.Size(143, 94)
        Me.Menu.Text = "Start Service"
        '
        'StartServiceToolStripMenuItem
        '
        Me.StartServiceToolStripMenuItem.CheckOnClick = True
        Me.StartServiceToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.StartServiceToolStripMenuItem.Image = Global.KeyLogger.My.Resources.Resources.service
        Me.StartServiceToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.White
        Me.StartServiceToolStripMenuItem.Name = "StartServiceToolStripMenuItem"
        Me.StartServiceToolStripMenuItem.Size = New System.Drawing.Size(142, 26)
        Me.StartServiceToolStripMenuItem.Text = "Start Service"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(139, 6)
        '
        'StatusToolStripMenuItem
        '
        Me.StatusToolStripMenuItem.Name = "StatusToolStripMenuItem"
        Me.StatusToolStripMenuItem.Size = New System.Drawing.Size(142, 26)
        Me.StatusToolStripMenuItem.Text = "Status"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(139, 6)
        '
        'ExitServiceToolStripMenuItem
        '
        Me.ExitServiceToolStripMenuItem.Name = "ExitServiceToolStripMenuItem"
        Me.ExitServiceToolStripMenuItem.Size = New System.Drawing.Size(142, 26)
        Me.ExitServiceToolStripMenuItem.Text = "Exit Service"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(353, 22)
        Me.Controls.Add(Me.ProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "My Service"
        Me.Menu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents NotifyIcon1 As NotifyIcon
    Friend WithEvents Menu As ContextMenuStrip
    Friend WithEvents StartServiceToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripSeparator
    Friend WithEvents ExitServiceToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StatusToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
End Class
