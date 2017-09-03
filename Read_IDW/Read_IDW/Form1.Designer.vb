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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(209, 30)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Directory create and delete"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 48)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(209, 30)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Open file"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(12, 92)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(531, 30)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Get File info"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(12, 128)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox2.Size = New System.Drawing.Size(531, 76)
        Me.TextBox2.TabIndex = 6
        Me.TextBox2.Tag = "File info"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(12, 256)
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox3.Size = New System.Drawing.Size(189, 99)
        Me.TextBox3.TabIndex = 7
        Me.TextBox3.Tag = "File info"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(12, 224)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(189, 23)
        Me.Button6.TabIndex = 8
        Me.Button6.Text = "Read Iproperties via Appentice"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(231, 224)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(239, 23)
        Me.Button8.TabIndex = 10
        Me.Button8.Text = "Change Iproperty ""author"" via Apprentice"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(231, 256)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(143, 23)
        Me.Button3.TabIndex = 11
        Me.Button3.Text = "Copy ipt file"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(6, 19)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(195, 23)
        Me.Button5.TabIndex = 12
        Me.Button5.Text = "Set part properties (inventor must run)"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(6, 48)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(173, 20)
        Me.TextBox4.TabIndex = 13
        Me.TextBox4.Text = "Part nummer 123"
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(6, 74)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(173, 20)
        Me.TextBox5.TabIndex = 14
        Me.TextBox5.Text = "Part descripton456"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.Controls.Add(Me.TextBox5)
        Me.GroupBox1.Controls.Add(Me.TextBox4)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 361)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(231, 117)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Set Iproperties"
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(231, 308)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(195, 26)
        Me.Button7.TabIndex = 16
        Me.Button7.Text = "Create Iproperty ""Fan type"""
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(580, 544)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Button6 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Button7 As Button
End Class
