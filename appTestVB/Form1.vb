Imports System.Drawing.Text
Imports System.IO
Imports System.Security.Cryptography
Imports OfficeOpenXml

Public Class Form1

	Dim id As Integer

	Public Sub New()
		InitializeComponent()
		id = 0
		lstvData.FullRowSelect = True
	End Sub

	Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
		Dim x As MyItem
		'i++   -->   i = i + 1
		Dim description As String = txtDescription.Text
		id = id + 1
		Dim Price As Random = New Random()

		x = New MyItem(id, description, Math.Round(Price.NextDouble() * 1000, 2))
		lstItems.Items.Add(x.ToString())

		'ListView -- ListViewItems -- SubItems

		For i = 1 To 100
			Dim row As ListViewItem = New ListViewItem(x.Id)
			row.SubItems.Add(x.Description)
			row.SubItems.Add(x.Price)
			lstvData.Items.Add(row)
			x.Id = x.Id + 1
			x.Price = Math.Round(Price.NextDouble() * 1000, 2)

		Next
		UpdateLabel()
		UpdateTotal()
	End Sub

	Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
		If lstvData.SelectedItems.Count = 0 Then
			Return
		End If
		For Each item As ListViewItem In lstvData.SelectedItems
			lstvData.Items.Remove(item)
		Next
		UpdateLabel()
		UpdateTotal()
	End Sub
	Private Sub btncreateexcel_Click(sender As Object, e As EventArgs) Handles btncreateexcel.Click

		Dim folderBrowserDialog As New FolderBrowserDialog()
		folderBrowserDialog.Description = "Select the folder"
		folderBrowserDialog.ShowNewFolderButton = False
		folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop

		If folderBrowserDialog.ShowDialog() <> DialogResult.OK Then
			Return
		End If

		Dim path As String = folderBrowserDialog.SelectedPath & "\Data.xlsx"

		Using package As New ExcelPackage()
				Dim worksheet = package.Workbook.Worksheets.Add("List")
				worksheet.Cells(1, 1).Value = "Id"
				worksheet.Cells(1, 2).Value = "Description"
				worksheet.Cells(1, 3).Value = "Price"
				Dim row As Integer = 2
			For Each item As ListViewItem In lstvData.Items
				worksheet.Cells(row, 1).Value = item.SubItems(0).Text
				worksheet.Cells(row, 2).Value = item.SubItems(1).Text
				worksheet.Cells(row, 3).Value = item.SubItems(2).Text
				row += 1
			Next
			package.SaveAs(New FileInfo(path))
			End Using

	End Sub

	Sub UpdateLabel()
		lblCount.Text = lstvData.Items.Count
	End Sub
	Sub UpdateTotal()
		Dim Total As Decimal = 0
		For Each item As ListViewItem In lstvData.Items
			Total = Total + Decimal.Parse(item.SubItems(2).Text)
		Next
		LblTotal.Text = Total
	End Sub


End Class
