Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Security.Authentication.ExtendedProtection

Public Class Form1
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dtOrigenOrigen As DataTable = Nothing

        Try
            If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
                Dim txtRuta = FolderBrowserDialog1.SelectedPath
                Debug.WriteLine(txtRuta)

                Dim ApExel = New Excel.Application
                Dim nBook As Excel.Workbook
                Dim nSheet As Excel.Worksheet

                nBook = ApExel.Workbooks.Add

                If ApExel.Application.Sheets.Count < 1 Then
                    nSheet = CType(nBook.Worksheets.Add, Excel.Worksheet)
                Else
                    nSheet = ApExel.Worksheets(1)
                    nSheet.Name = "Hoja1"
                End If

                Dim oExel = New Excel.Application
                Dim ApExel_TitleSet As Boolean = False

                Dim columna As Integer = 2

                For Each files As String In Directory.GetFiles(txtRuta)

                    Dim strcamino As String = files
                    oExel.Workbooks.Open(strcamino)
                    Dim oBook As Object
                    Dim oSheet As Object
                    oBook = oExel.ActiveWorkbook
                    oSheet = oExel.Worksheets(2)

                    Debug.WriteLine(files)
                    Debug.WriteLine(Path.GetFileNameWithoutExtension(files))
                    Debug.WriteLine(Path.GetFileNameWithoutExtension(files).Substring(6, 8))
                    Dim fecha = Path.GetFileNameWithoutExtension(files).Substring(6, 8)

                    Debug.WriteLine(fecha.Substring(0, 4))
                    Debug.WriteLine(fecha.Substring(4, 2))
                    Debug.WriteLine(fecha.Substring(6, 2))
                    Dim ano = fecha.Substring(0, 4)
                    Dim mes = fecha.Substring(4, 2)
                    Dim dia = fecha.Substring(6, 2)


                    If nSheet.Name = "Hoja1" Then
                        nSheet.Name = "G" & mes
                    End If

                    Debug.WriteLine("Sheet Name: " & nSheet.Name.Substring(1, 2))
                    Debug.WriteLine("Sheet Count: " & ApExel.Application.Sheets.Count)
                    If mes > nSheet.Name.Substring(1, 2) And dia = 1 Then
                        nSheet = CType(nBook.Worksheets.Add, Excel.Worksheet)
                        'nSheet = ApExel.Worksheets(nSheet.Count)
                        columna = 2
                        nSheet = nBook.Worksheets(ApExel.Application.Sheets.Count)
                        nSheet.Name = "G" & mes
                        ApExel_TitleSet = False
                    End If

                    'Copia de la columna de titulos
                    If ApExel_TitleSet = False Then
                        oSheet.Range("A4:A25").Copy()

                        'Dim titulo = oSheet.Sheets(1).Range("A1").Copy()
                        nSheet.Range("A1").PasteSpecial(Excel.XlPasteType.xlPasteValues)
                        With nSheet.Columns("A")
                            .ColumnWidth = 36
                        End With
                        'nSheet.Columns("A").EntireColumn.AutoFit()
                        nSheet.Range("A23").Value = "Fecha"
                        nSheet.Range("A23").Font.Name = "Arial"
                        nSheet.Range("A23").Font.Size = 8

                        ApExel_TitleSet = True
                    End If

                    'Copia de columna de information
                    oSheet.Range("C4:C25").Copy()
                    nSheet.Cells(1, columna).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                    With nSheet.Columns(columna)
                        .ColumnWidth = 15
                    End With

                    'Fotmateo de la fecha
                    Dim fechaFinal As String = dia & "-" & mes & "-" & ano
                    nSheet.Cells(23, columna).Value = DateSerial(Month:=mes, Day:=dia, Year:=ano)
                    nSheet.Cells(23, columna).NumberFormat = "dd-mmm"
                    nSheet.Cells(23, columna).font.Name = "Arial"
                    nSheet.Cells(23, columna).font.Size = 8

                    columna += 1

                    nSheet.Move(After:=nBook.Worksheets(nBook.Worksheets.Count))

                    oBook.Close()
                    oBook = Nothing

                Next

                SaveFileDialog1.DefaultExt = "*.xlsx"
                SaveFileDialog1.FileName = "PrecioElectricidad"
                SaveFileDialog1.Filter = "Archivos de Exel (*.xlsx)|*.xlsx"
                SaveFileDialog1.ShowDialog()

                nBook.SaveAs(SaveFileDialog1.FileName)
                MessageBox.Show("El archivo se creo y se guardo en:" & vbCrLf & SaveFileDialog1.FileName)

                nBook.Close()
                nBook = Nothing
                ApExel = Nothing

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fileTest As String = "C:\Users\Davidf\Desktop\test.xlsx"
        If File.Exists(fileTest) Then
            File.Delete(fileTest)
        End If

        Dim oExcel As Object
        oExcel = CreateObject("Excel.Application")
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        oBook = oExcel.Workbooks.Add

        ' first worksheet
        If oExcel.Application.Sheets.Count() < 1 Then
            oSheet = CType(oBook.Worksheets.Add(), Excel.Worksheet)
        Else
            oSheet = oExcel.Worksheets(1)
        End If
        oSheet.Name = "one"
        oSheet.Range("B1").Value = "First One"

        oSheet.Move(After:=oBook.Worksheets(oBook.Worksheets.Count))

        ' second
        If oExcel.Application.Sheets.Count() < 2 Then
            oSheet = CType(oBook.Worksheets.Add(), Excel.Worksheet)
        Else
            oSheet = oExcel.Worksheets(2)
        End If
        oSheet.Name = "two"
        oSheet.Range("B1").Value = "Second one"

        oSheet.Move(After:=oBook.Worksheets(oBook.Worksheets.Count))

        ' third
        If oExcel.Application.Sheets.Count() < 3 Then
            oSheet = CType(oBook.Worksheets.Add(), Excel.Worksheet)
        Else
            oSheet = oExcel.Worksheets(3)
        End If
        oSheet.Name = "three"
        oSheet.Range("B1").Value = "Thrid"

        oSheet.Move(After:=oBook.Worksheets(oBook.Worksheets.Count))

        ' next
        If oExcel.Application.Sheets.Count() < 4 Then
            oSheet = CType(oBook.Worksheets.Add(), Excel.Worksheet)
        Else
            oSheet = oExcel.Worksheets(4)
        End If
        oSheet.Name = "four"
        oSheet.Range("B1").Value = "Four"

        ' by default this worksheet will be placed in front of the first
        ' the code below will move it after third one

        oSheet.Move(After:=oBook.Worksheets(oBook.Worksheets.Count))

        oBook.SaveAs(fileTest)
        oBook.Close()
        oBook = Nothing

        MessageBox.Show("Fin")

    End Sub
End Class
