<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form10
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
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtLatLon = New System.Windows.Forms.TextBox()
        Me.btnGeocode = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(15, 25)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(356, 20)
        Me.txtAddress.TabIndex = 0
        '
        'txtLatLon
        '
        Me.txtLatLon.Location = New System.Drawing.Point(15, 64)
        Me.txtLatLon.Name = "txtLatLon"
        Me.txtLatLon.Size = New System.Drawing.Size(356, 20)
        Me.txtLatLon.TabIndex = 1
        '
        'btnGeocode
        '
        Me.btnGeocode.Location = New System.Drawing.Point(377, 22)
        Me.btnGeocode.Name = "btnGeocode"
        Me.btnGeocode.Size = New System.Drawing.Size(75, 23)
        Me.btnGeocode.TabIndex = 2
        Me.btnGeocode.Text = "Geocode it"
        Me.btnGeocode.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Enter Address"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Result"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(460, 97)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnGeocode)
        Me.Controls.Add(Me.txtLatLon)
        Me.Controls.Add(Me.txtAddress)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtLatLon As System.Windows.Forms.TextBox
    Friend WithEvents btnGeocode As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
