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
        Me.dgvDados = New System.Windows.Forms.DataGridView()
        Me.btnEncerrar = New System.Windows.Forms.Button()
        Me.btnExportar = New System.Windows.Forms.Button()
        CType(Me.dgvDados, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDados
        '
        Me.dgvDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDados.Location = New System.Drawing.Point(12, 12)
        Me.dgvDados.Name = "dgvDados"
        Me.dgvDados.Size = New System.Drawing.Size(582, 235)
        Me.dgvDados.TabIndex = 0
        '
        'btnEncerrar
        '
        Me.btnEncerrar.Location = New System.Drawing.Point(12, 253)
        Me.btnEncerrar.Name = "btnEncerrar"
        Me.btnEncerrar.Size = New System.Drawing.Size(102, 59)
        Me.btnEncerrar.TabIndex = 1
        Me.btnEncerrar.Text = "Encerrar"
        Me.btnEncerrar.UseVisualStyleBackColor = True
        '
        'btnExportar
        '
        Me.btnExportar.Location = New System.Drawing.Point(492, 253)
        Me.btnExportar.Name = "btnExportar"
        Me.btnExportar.Size = New System.Drawing.Size(102, 59)
        Me.btnExportar.TabIndex = 2
        Me.btnExportar.Text = "Excel"
        Me.btnExportar.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(606, 324)
        Me.Controls.Add(Me.btnExportar)
        Me.Controls.Add(Me.btnEncerrar)
        Me.Controls.Add(Me.dgvDados)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.dgvDados, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDados As System.Windows.Forms.DataGridView
    Friend WithEvents btnEncerrar As System.Windows.Forms.Button
    Friend WithEvents btnExportar As System.Windows.Forms.Button

End Class
