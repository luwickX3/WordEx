Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Dialogo = New SaveFileDialog()
        If Dialogo.ShowDialog() <> DialogResult.OK Then
            End
        End If
        Dim Ruta = Dialogo.FileName
        Dim WordApp = New Microsoft.Office.Interop.Word.Application
        Dim WordDoc = WordApp.Documents.Add()
        WordApp.Selection.TypeText(TextBox1.Text)
        WordApp.ActiveDocument.SaveAs2(Ruta)
        WordApp.Visible = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Dialogo = New SaveFileDialog()
        If Dialogo.ShowDialog() <> DialogResult.OK Then
            End
        End If
        Dim Ruta = Dialogo.FileName
        Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
        Dim ExcelBook = ExcelApp.Workbooks.Add()
        ExcelBook.Sheets(1).Cells(1, 1) = TextBox1.Text
        ExcelBook.SaveAs(Ruta)
        ExcelApp.Visible = True
    End Sub
End Class
