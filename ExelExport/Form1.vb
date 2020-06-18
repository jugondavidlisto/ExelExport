Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Security.Authentication.ExtendedProtection

Public Class Form1
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Variables Exel Export
        Dim exportExel = New Excel.Application
        Dim exportBook As Excel.Workbook
        Dim exportSheet As Excel.Worksheet
        Dim ApExel_TitleSet As Boolean = False
        Dim columna As Integer = 2

        'Creacion de la Hoja de Exel Export y seteo del nombre por defecto
        exportBook = exportExel.Workbooks.Add

        If exportExel.Application.Sheets.Count < 1 Then
            exportSheet = CType(exportBook.Worksheets.Add, Excel.Worksheet)
        Else
            exportSheet = exportExel.Worksheets(1)
            exportSheet.Name = "Hoja1"
        End If

        'Variables Exel Original
        Dim originalExel = New Excel.Application
        Dim originalBook As Object
        Dim originalSheet As Object

        'Funcion principal
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Dim txtRuta = FolderBrowserDialog1.SelectedPath
            Debug.WriteLine(txtRuta)

            'Variables Barra de Progreso
            Dim Cuenta As Integer = 0
            ProgressBar1.Maximum = Directory.GetFiles(txtRuta).Count
            Label1.Visible = False
            Label2.Text = 0 & (" %")

            Try
                For Each files As String In Directory.GetFiles(txtRuta)

                    Dim strcamino As String = files
                    originalExel.Workbooks.Open(strcamino)

                    originalBook = originalExel.ActiveWorkbook
                    originalSheet = originalExel.Worksheets(2)

                    ' Obtencion de la fecha actual apartir de los nombres de archivo
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

                    ' Correcion del nombre de la Hoja del Exel Export por el formato correcto G+numeroMes
                    If exportSheet.Name = "Hoja1" Then
                        exportSheet.Name = "G" & mes
                    End If

                    ' Condicional para la creacion de una nueva hoja si a terminado de llenar un mes
                    If mes > exportSheet.Name.Substring(1, 2) And dia = 1 Then
                        exportSheet = CType(exportBook.Worksheets.Add(), Excel.Worksheet)
                        'nSheet = ApExel.Worksheets(nSheet.Count)
                        columna = 2
                        'exportSheet = exportBook.Worksheets(exportExel.Application.Sheets.Count)
                        exportSheet.Name = "G" & mes
                        ApExel_TitleSet = False
                    End If
                    Debug.WriteLine("Sheet Name: " & exportSheet.Name.Substring(1, 2))
                    Debug.WriteLine("Sheet Count: " & exportExel.Application.Sheets.Count)

                    'Copia de la columna de titulos
                    If ApExel_TitleSet = False Then
                        originalSheet.Range("A4:A25").Copy()

                        'Dim titulo = oSheet.Sheets(1).Range("A1").Copy()
                        exportSheet.Range("A1").PasteSpecial(Excel.XlPasteType.xlPasteValues)
                        With exportSheet.Columns("A")
                            .ColumnWidth = 36
                        End With
                        'nSheet.Columns("A").EntireColumn.AutoFit()
                        exportSheet.Range("A23").Value = "Fecha"
                        exportSheet.Range("A23").Font.Name = "Arial"
                        exportSheet.Range("A23").Font.Size = 8

                        ApExel_TitleSet = True
                    End If

                    'Copia de columna de information
                    originalSheet.Range("C4:C25").Copy()
                    exportSheet.Cells(1, columna).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                    With exportSheet.Columns(columna)
                        .ColumnWidth = 15
                    End With

                    'Fotmateo de la fecha
                    Dim fechaFinal As String = dia & "-" & mes & "-" & ano
                    exportSheet.Cells(23, columna).Value = DateSerial(Month:=mes, Day:=dia, Year:=ano)
                    exportSheet.Cells(23, columna).NumberFormat = "dd-mmm"
                    exportSheet.Cells(23, columna).font.Name = "Arial"
                    exportSheet.Cells(23, columna).font.Size = 8

                    columna += 1

                    ' por defecto esta hoja de trabajo se colocará delante de la primera
                    ' el siguiente código lo moverá atras la hoja actual
                    exportSheet.Move(After:=exportBook.Worksheets(exportBook.Worksheets.Count))

                    ' Cierre del libro original
                    originalBook.Close()
                    originalBook = Nothing

                    ' Variables ProgressBar
                    ProgressBar1.Value = Cuenta
                    Cuenta = Cuenta + 1
                    Label2.Text = CLng((ProgressBar1.Value * 100) / ProgressBar1.Maximum) & " %"
                Next

                Label1.Visible = True

                ' Guardado del exel export y cerrar y vaciar variables
                SaveFileDialog1.DefaultExt = "*.xlsx"
                SaveFileDialog1.FileName = "PrecioElectricidad"
                SaveFileDialog1.Filter = "Archivos de Exel (*.xlsx)|*.xlsx"
                SaveFileDialog1.ShowDialog()

                exportBook.SaveAs(SaveFileDialog1.FileName)
                'MessageBox.Show("El archivo se creo y se guardo en:" & vbCrLf & SaveFileDialog1.FileName)

                exportBook.Close()
                exportBook = Nothing
                exportExel = Nothing
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

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

        For number As Double = 1 To 4 Step +1

            If oExcel.Application.Sheets.Count() < number Then
                oSheet = CType(oBook.Worksheets.Add(), Excel.Worksheet)
            Else
                oSheet = oExcel.Worksheets(number)
            End If
            oSheet.Name = number
            oSheet.Range("B1").Value = number

            ' por defecto esta hoja de trabajo se colocará delante de la primera
            ' el siguiente código lo moverá atras la hoja actual

            oSheet.Move(After:=oBook.Worksheets(oBook.Worksheets.Count))

        Next

        oBook.SaveAs(fileTest)
        oBook.Close()
        oBook = Nothing

        MessageBox.Show("Fin")

    End Sub
End Class
