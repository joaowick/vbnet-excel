Imports Microsoft.Office.Interop

Public Class exportExcel
    Dim XcelApp As New Excel.Application()

    Private Sub carregaGrid()
        Try
            Dim dt As New DataTable
            dt.Columns.Add("Codigo", GetType(Integer))
            dt.Columns.Add("Nome", GetType(String))
            dt.Columns.Add("Admissao", GetType(DateTime))
            dt.Columns.Add("Setor", GetType(Integer))
            dt.Columns.Add("Salario", GetType(Double))
            Dim dr As DataRow = dt.NewRow()
            dr("Codigo") = 1
            dr("Nome") = "João Torres"
            dr("Admissao") = DateTime.Now
            dr("Setor") = 20
            dr("Salario") = 20000
            dt.Rows.Add(dr)
            dr = dt.NewRow()
            dr("Codigo") = 2
            dr("Nome") = "Jennifer"
            dr("Admissao") = DateTime.Now
            dr("Setor") = 40
            dr("Salario") = 20000
            dt.Rows.Add(dr)
            dgvDados.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("Erro" + ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        carregaGrid()
    End Sub

    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        If dgvDados.Rows.Count > 0 Then
            Try
                XcelApp.Application.Workbooks.Add(Type.Missing)

                For i As Integer = 1 To dgvDados.Columns.Count
                    XcelApp.Cells(1, i) = dgvDados.Columns(i - 1).HeaderText
                Next
                '
                For i As Integer = 0 To dgvDados.Rows.Count - 2
                    For j As Integer = 0 To dgvDados.Columns.Count - 1
                        XcelApp.Cells(i + 2, j + 1) = dgvDados.Rows(i).Cells(j).Value.ToString()
                    Next
                Next
                '
                XcelApp.Columns.AutoFit()
                '
                XcelApp.Visible = True
            Catch ex As Exception
                MessageBox.Show("Erro: " + ex.Message)
                XcelApp.Quit()
            End Try
        End If
    End Sub

    Private Sub btnEncerrar_Click(sender As Object, e As EventArgs) Handles btnEncerrar.Click
        XcelApp.Quit()
        XcelApp = Nothing
        Me.Close()
    End Sub

    Private Sub dgvDados_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDados.CellContentClick

    End Sub
End Class
