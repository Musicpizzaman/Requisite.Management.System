<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Newreq
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.rtbDiscription = New System.Windows.Forms.RichTextBox()
        Me.cboEntity = New System.Windows.Forms.ComboBox()
        Me.cboQFD = New System.Windows.Forms.ComboBox()
        Me.cboTime = New System.Windows.Forms.ComboBox()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 365)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(295, 50)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "&Submit"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(12, 12)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(141, 24)
        Me.txtName.TabIndex = 1
        Me.txtName.Text = "Trait Name"
        '
        'rtbDiscription
        '
        Me.rtbDiscription.Location = New System.Drawing.Point(12, 106)
        Me.rtbDiscription.Name = "rtbDiscription"
        Me.rtbDiscription.Size = New System.Drawing.Size(294, 253)
        Me.rtbDiscription.TabIndex = 5
        Me.rtbDiscription.Text = "Enter Description"
        '
        'cboEntity
        '
        Me.cboEntity.AllowDrop = True
        Me.cboEntity.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cboEntity.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboEntity.FormattingEnabled = True
        Me.cboEntity.Items.AddRange(New Object() {"User", "Safety", "Constraint", "Quality "})
        Me.cboEntity.Location = New System.Drawing.Point(12, 42)
        Me.cboEntity.Name = "cboEntity"
        Me.cboEntity.Size = New System.Drawing.Size(141, 26)
        Me.cboEntity.TabIndex = 4
        Me.cboEntity.Text = "Entity Type"
        '
        'cboQFD
        '
        Me.cboQFD.AllowDrop = True
        Me.cboQFD.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cboQFD.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboQFD.FormattingEnabled = True
        Me.cboQFD.Items.AddRange(New Object() {"Expected", "Normal", "Exciting"})
        Me.cboQFD.Location = New System.Drawing.Point(173, 42)
        Me.cboQFD.Name = "cboQFD"
        Me.cboQFD.Size = New System.Drawing.Size(134, 26)
        Me.cboQFD.TabIndex = 6
        Me.cboQFD.Text = "QFD"
        '
        'cboTime
        '
        Me.cboTime.AllowDrop = True
        Me.cboTime.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cboTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTime.FormattingEnabled = True
        Me.cboTime.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"})
        Me.cboTime.Location = New System.Drawing.Point(12, 74)
        Me.cboTime.Name = "cboTime"
        Me.cboTime.Size = New System.Drawing.Size(295, 26)
        Me.cboTime.TabIndex = 7
        Me.cboTime.Text = "Expected Completion in Weeks"
        '
        'txtPath
        '
        Me.txtPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPath.Location = New System.Drawing.Point(173, 12)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(134, 24)
        Me.txtPath.TabIndex = 8
        Me.txtPath.Text = "Path"
        '
        'Newreq
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(317, 425)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.cboTime)
        Me.Controls.Add(Me.cboQFD)
        Me.Controls.Add(Me.rtbDiscription)
        Me.Controls.Add(Me.cboEntity)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Newreq"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents rtbDiscription As System.Windows.Forms.RichTextBox
    Friend WithEvents cboEntity As System.Windows.Forms.ComboBox
    Friend WithEvents cboQFD As System.Windows.Forms.ComboBox
    Friend WithEvents cboTime As System.Windows.Forms.ComboBox
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
End Class
