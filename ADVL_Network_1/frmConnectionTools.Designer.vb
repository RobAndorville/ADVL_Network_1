<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConnectionTools
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnUpdateStatus = New System.Windows.Forms.Button()
        Me.txtConnectionStatus = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtMessageServiceStatus = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnStopMsgService = New System.Windows.Forms.Button()
        Me.btnStartMsgService = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(925, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnUpdateStatus)
        Me.GroupBox1.Controls.Add(Me.txtConnectionStatus)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtMessageServiceStatus)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 52)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(318, 114)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Status:"
        '
        'btnUpdateStatus
        '
        Me.btnUpdateStatus.Location = New System.Drawing.Point(6, 78)
        Me.btnUpdateStatus.Name = "btnUpdateStatus"
        Me.btnUpdateStatus.Size = New System.Drawing.Size(64, 22)
        Me.btnUpdateStatus.TabIndex = 9
        Me.btnUpdateStatus.Text = "Update"
        Me.btnUpdateStatus.UseVisualStyleBackColor = True
        '
        'txtConnectionStatus
        '
        Me.txtConnectionStatus.Location = New System.Drawing.Point(104, 49)
        Me.txtConnectionStatus.Name = "txtConnectionStatus"
        Me.txtConnectionStatus.ReadOnly = True
        Me.txtConnectionStatus.Size = New System.Drawing.Size(208, 20)
        Me.txtConnectionStatus.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Connection:"
        '
        'txtMessageServiceStatus
        '
        Me.txtMessageServiceStatus.Location = New System.Drawing.Point(104, 23)
        Me.txtMessageServiceStatus.Name = "txtMessageServiceStatus"
        Me.txtMessageServiceStatus.ReadOnly = True
        Me.txtMessageServiceStatus.Size = New System.Drawing.Size(208, 20)
        Me.txtMessageServiceStatus.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Message Service:"
        '
        'btnStopMsgService
        '
        Me.btnStopMsgService.Location = New System.Drawing.Point(12, 172)
        Me.btnStopMsgService.Name = "btnStopMsgService"
        Me.btnStopMsgService.Size = New System.Drawing.Size(132, 22)
        Me.btnStopMsgService.TabIndex = 10
        Me.btnStopMsgService.Text = "Stop Message Service"
        Me.btnStopMsgService.UseVisualStyleBackColor = True
        '
        'btnStartMsgService
        '
        Me.btnStartMsgService.Location = New System.Drawing.Point(12, 200)
        Me.btnStartMsgService.Name = "btnStartMsgService"
        Me.btnStartMsgService.Size = New System.Drawing.Size(132, 22)
        Me.btnStartMsgService.TabIndex = 11
        Me.btnStartMsgService.Text = "Start Message Service"
        Me.btnStartMsgService.UseVisualStyleBackColor = True
        '
        'frmConnectionTools
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(985, 843)
        Me.Controls.Add(Me.btnStartMsgService)
        Me.Controls.Add(Me.btnStopMsgService)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmConnectionTools"
        Me.Text = "Message Service Connection Tools"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnUpdateStatus As Button
    Friend WithEvents txtConnectionStatus As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtMessageServiceStatus As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnStopMsgService As Button
    Friend WithEvents btnStartMsgService As Button
End Class
