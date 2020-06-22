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
        Dim originalBook As Excel.Workbook
        Dim originalSheet As Excel.Worksheet

        'Funcion principal
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Dim txtRuta = FolderBrowserDialog1.SelectedPath
            Debug.WriteLine(txtRuta)
            TextBox1.Text = txtRuta
            Button2.Enabled = False

            'Variables Barra de Progreso
            Dim Cuenta As Integer = 0
            ProgressBar1.Value = Cuenta
            Debug.WriteLine(Directory.GetFiles(txtRuta, "*.xls").Count)
            ProgressBar1.Maximum = Directory.GetFiles(txtRuta, "*.xls").Count
            Label1.Visible = False
            Label2.Text = 0 & (" %")

            'Variable del tiempo inicial
            Dim starttime As DateTime = DateTime.Now

            Try
                For Each files As String In Directory.GetFiles(txtRuta)

                    'Calcula el lapso de tiempo entre el ahora y el tiempo antes de ejecutar el bucle
                    Dim timespent As TimeSpan = DateTime.Now - starttime

                    'Filtro de archivos solo .xls	
                    Debug.WriteLine(Path.GetExtension(files) & " " & files)
                    If Path.GetExtension(files) = ".xls" Then
                        Dim strcamino As String = files
                        originalExel.Workbooks.Open(strcamino)
                        originalBook = originalExel.ActiveWorkbook
                        originalSheet = originalExel.Worksheets(2)

                        ' Obtencion de la fecha actual apartir de los nombres de archivo
                        Debug.WriteLine(files)
                        Label1.Text = "Procesando archivo " & Path.GetFileNameWithoutExtension(files)
                        Label1.Visible = True
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
                            Threading.Thread.Sleep(1000)
                            'Dim titulo = oSheet.Sheets(1).Range("A1").Copy()
                            exportSheet.Range("A1:A22").PasteSpecial(Excel.XlPasteType.xlPasteValues)
                            'Threading.Thread.Sleep(500)
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
                        Threading.Thread.Sleep(1000)
                        exportSheet.Cells(1, columna).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                        'Threading.Thread.Sleep(500)
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
                        'Threading.Thread.Sleep(500)

                        ' Cierre del libro original
                        originalBook.Close(False)
                        ReleaseComObject(originalBook)
                        'originalBook = Nothing

                        ' Variables ProgressBar
                        Cuenta = Cuenta + 1
                        ProgressBar1.Value = Cuenta
                        Label2.Text = CLng((ProgressBar1.Value * 100) / ProgressBar1.Maximum) & " %"

                        ' Funcion de conversion de tiempo y mostrar al usuario
                        Dim secondsremaining As Integer = CInt((timespent.TotalSeconds / ProgressBar1.Value * (ProgressBar1.Maximum - ProgressBar1.Value)))

                        Debug.WriteLine(secondsremaining)

                        Dim Hours = secondsremaining / 3600
                        secondsremaining = secondsremaining Mod 3600
                        Dim Minutes = secondsremaining / 60
                        secondsremaining = secondsremaining Mod 60

                        Debug.WriteLine(Hours & " h" & Minutes & " min" & secondsremaining & " sec")

                        If (Hours.ToString().Substring(0, 1) <> 0) Then
                            Label4.Text = "Tiempo estimado restante: " & Hours.ToString().Substring(0, 1) & " h" & Minutes.ToString().Substring(0, 1) & " mi n" & Math.Round(secondsremaining) & " sec"
                            Label4.Visible = True
                        ElseIf (Minutes.ToString.Substring(0, 1) <> 0) Then
                            Label4.Text = "Tiempo estimado restante: " & Minutes.ToString().Substring(0, 1) & " min " & Math.Round(secondsremaining) & " sec"
                            Label4.Visible = True
                        ElseIf (Math.Round(secondsremaining) <> 0) Then
                            Label4.Text = "Tiempo estimado restante:  " & Math.Round(secondsremaining) & " sec"
                            Label4.Visible = True
                        End If

                        Label1.Visible = False

                        'Limpiar referencias o eso creo xD
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End If
                Next
                Label4.Visible = False
                Label1.Text = "Finalizado"
                Label1.Visible = True

                ' Guardado del exel export y cerrar y vaciar variables
                SaveFileDialog1.DefaultExt = "*.xlsx"
                SaveFileDialog1.FileName = "PrecioElectricidad"
                SaveFileDialog1.Filter = "Archivos de Exel (*.xlsx)|*.xlsx"
                SaveFileDialog1.ShowDialog()

                exportBook.SaveAs(SaveFileDialog1.FileName)
                'MessageBox.Show("El archivo se creo y se guardo en:" & vbCrLf & SaveFileDialog1.FileName)

                '~~> Close the File
                exportBook.Close(False)

                '~~> Quit the Excel Application
                exportExel.Quit()

                '~~> Clean Up
                ReleaseComObject(exportExel)
                ReleaseComObject(exportBook)
                ReleaseComObject(exportSheet)

                Button2.Enabled = True
            Catch ex As Exception
                Label4.Visible = False
                MessageBox.Show(ex.Message)

                SaveFileDialog1.DefaultExt = "*.xlsx"
                SaveFileDialog1.FileName = "PrecioElectricidad"
                SaveFileDialog1.Filter = "Archivos de Exel (*.xlsx)|*.xlsx"
                SaveFileDialog1.ShowDialog()

                exportBook.SaveAs(SaveFileDialog1.FileName)

                '~~> Close the File
                exportBook.Close(False)

                '~~> Quit the Excel Application
                exportExel.Quit()

                '~~> Clean Up
                ReleaseComObject(exportExel)
                ReleaseComObject(exportBook)
                ReleaseComObject(exportSheet)

                If (originalBook Is Nothing) Then
                    '~~> Quit the Excel Application
                    originalExel.Quit()

                    '~~> Clean Up
                    ReleaseComObject(originalExel)
                Else
                    '~~> Close the File
                    originalBook.Close(False)

                    '~~> Quit the Excel Application
                    originalExel.Quit()

                    '~~> Clean Up
                    ReleaseComObject(originalBook)
                    ReleaseComObject(originalExel)
                    ReleaseComObject(originalSheet)
                End If

                End
            End Try
        End If

    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
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
