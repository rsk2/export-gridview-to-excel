Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.IO

Public Class TestPage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub cmdExport_Click(sender As Object, e As System.EventArgs) Handles cmdExport.Click
        'You need an Export folder inside your project
        Dim directoryName As String = Server.MapPath("~/Export/")
        Dim localExcelPath As String = directoryName & lblTitle.Text & ".xlsx"
        'Just in case there is a file existing with same name
        localExcelPath = CheckAndRenameIfFileExists(localExcelPath)

        Dim excel = New Application()
        excel.Visible = True
        Dim wb As Workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        Dim sh As Worksheet = wb.Sheets.Add()
        sh.Name = "Scorecard"
        Dim tb As New Table()
        Dim rowCount As Integer = 1
        AddGridToWorksheet(gvTest, tb, sh, rowCount)
        sh.Columns.AutoFit()
        wb.SaveAs(localExcelPath)
        wb.Close(True)
        excel.Quit()
        Response.Clear()
        Response.Charset = ""
        Response.Buffer = True
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        Response.AddHeader("content-disposition", "attachment;filename=" & lblTitle.Text & ".xlsx")
        Response.TransmitFile((localExcelPath))
        Response.End()
    End Sub

    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    Private Function CheckAndRenameIfFileExists(fullPath As String) As String
        Dim count As Integer = 0
        Dim fileNameOnly As String = Path.GetFileNameWithoutExtension(fullPath)
        Dim extension As String = Path.GetExtension(fullPath)
        Dim directory As String = Path.GetDirectoryName(fullPath)
        Dim newFullPath As String = fullPath
        Dim tempFileName As String
        While (File.Exists(newFullPath))
            count += 1
            tempFileName = String.Format("{0}({1})", fileNameOnly, count)
            newFullPath = Path.Combine(directory, tempFileName + extension)
        End While
        Return newFullPath
    End Function

    Public Sub AddGridToWorksheet(gvCtrl As GridView, tb As Table, sh As Worksheet, ByRef rowCount As Integer)
        If gvCtrl.Visible = True Then
            Dim j As Integer = 1
            For Each cell As TableCell In gvCtrl.HeaderRow.Cells
                If cell.Text <> "&nbsp;" Then
                    sh.Cells(rowCount, GetExcelColumnName(j)).Value = cell.Text
                    j = j + 1
                End If
            Next
            rowCount += 1

            For Each row As GridViewRow In gvCtrl.Rows
                sh.Cells(rowCount, GetExcelColumnName(1)).Value = row.Cells(0).Text
                sh.Cells(rowCount, GetExcelColumnName(2)).Value = row.Cells(1).Text
                sh.Cells(rowCount, GetExcelColumnName(3)).Value = row.Cells(2).Text
                sh.Cells(rowCount, GetExcelColumnName(4)).Value = row.Cells(3).Text
                sh.Cells(rowCount, GetExcelColumnName(5)).Value = row.Cells(4).Text
                sh.Cells(rowCount, GetExcelColumnName(6)).Value = row.Cells(5).Text
                sh.Cells(rowCount, GetExcelColumnName(7)).Value = row.Cells(6).Text
                rowCount += 1
            Next
        End If
    End Sub

End Class