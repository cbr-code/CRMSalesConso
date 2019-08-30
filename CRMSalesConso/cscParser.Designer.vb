<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class cscParser
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(cscParser))
        Me.dgvSales = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.dgvTrans = New System.Windows.Forms.DataGridView()
        Me.dgvpayments = New System.Windows.Forms.DataGridView()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        CType(Me.dgvSales, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvTrans, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvpayments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSales
        '
        Me.dgvSales.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSales.Location = New System.Drawing.Point(12, 209)
        Me.dgvSales.Name = "dgvSales"
        Me.dgvSales.Size = New System.Drawing.Size(776, 55)
        Me.dgvSales.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 508)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(78, 31)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dgvTrans
        '
        Me.dgvTrans.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTrans.Location = New System.Drawing.Point(12, 312)
        Me.dgvTrans.Name = "dgvTrans"
        Me.dgvTrans.Size = New System.Drawing.Size(776, 39)
        Me.dgvTrans.TabIndex = 2
        '
        'dgvpayments
        '
        Me.dgvpayments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvpayments.Location = New System.Drawing.Point(12, 384)
        Me.dgvpayments.Name = "dgvpayments"
        Me.dgvpayments.Size = New System.Drawing.Size(776, 53)
        Me.dgvpayments.TabIndex = 3
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Inv Conso Trans"
        Me.NotifyIcon1.Visible = True
        '
        'cscParser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(658, 133)
        Me.ControlBox = False
        Me.Controls.Add(Me.dgvpayments)
        Me.Controls.Add(Me.dgvTrans)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dgvSales)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "cscParser"
        Me.Text = "cscParser"
        CType(Me.dgvSales, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvTrans, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvpayments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgvSales As DataGridView
    Friend WithEvents Button1 As Button
    Friend WithEvents dgvTrans As DataGridView
    Friend WithEvents dgvpayments As DataGridView
    Friend WithEvents NotifyIcon1 As NotifyIcon
End Class
