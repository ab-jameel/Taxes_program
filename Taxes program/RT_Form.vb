Imports Microsoft.Office.Interop.excel
Imports System.Data.OleDb
Public Class RT_Form
    Dim dtTable2 As New Data.DataTable
    Dim da2 As New OleDbDataAdapter
    Dim exlapp As New Application
    Dim exlworkbook As Workbook
    Dim exlworksheet As Worksheet
    Dim missvalue As Object = System.Reflection.Missing.Value

    Sub excel1()
        If CDbl(Label9.Text) <> 0 And CDbl(Label6.Text) <> 0 Then
            Try
                exlworkbook = exlapp.Workbooks.Add(missvalue)

                Try
                    ' For English Office
                    exlworksheet = exlworkbook.Sheets("sheet1")

                    'Buys
                    RB_Form.DateTimePicker1.Value = DateTimePicker1.Value
                    RB_Form.DateTimePicker2.Value = DateTimePicker2.Value.AddDays(1)
                    RB_Form.Show()
                    RB_Form.Visible = False

                    'sales
                    RS_Form.DateTimePicker1.Value = DateTimePicker1.Value
                    RS_Form.DateTimePicker2.Value = DateTimePicker2.Value.AddDays(1)
                    RS_Form.Show()
                    RS_Form.Visible = False

                    'Buys
                    For colh = 0 To RB_Form.DataGridView1.ColumnCount - 1
                        exlworksheet.Cells(5, colh + 2) = RB_Form.DataGridView1.Columns(colh).HeaderText + ":"
                    Next

                    'Sales
                    For colh1 = 0 To RS_Form.DataGridView1.ColumnCount - 1
                        exlworksheet.Cells(5, colh1 + 16) = RS_Form.DataGridView1.Columns(colh1).HeaderText + ":"
                    Next

                    'Buys
                    For i As Integer = 0 To RB_Form.DataGridView1.RowCount - 1
                        For j As Integer = 0 To RB_Form.DataGridView1.ColumnCount - 1
                            exlworksheet.Cells(i + 6, j + 2) = RB_Form.DataGridView1.Rows(i).Cells(j).Value.ToString
                        Next
                    Next

                    'Sales
                    For i1 As Integer = 0 To RS_Form.DataGridView1.RowCount - 1
                        For j1 As Integer = 0 To RS_Form.DataGridView1.ColumnCount - 1
                            exlworksheet.Cells(i1 + 6, j1 + 16) = RS_Form.DataGridView1.Rows(i1).Cells(j1).Value.ToString
                        Next
                    Next

                    'Buys
                    exlworksheet.Range("B1").EntireColumn.Delete()
                    exlworksheet.Range("D1").EntireColumn.Delete()
                    exlworksheet.Range("G1").EntireColumn.Delete()
                    exlworksheet.Range("B:B").Copy()
                    exlworksheet.Range("I1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("H:H").Copy()
                    exlworksheet.Range("B1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("C:C").Copy()
                    exlworksheet.Range("H1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("G:G").Copy()
                    exlworksheet.Range("C1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("E:E").Copy()
                    exlworksheet.Range("G1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("D:D").Copy()
                    exlworksheet.Range("E1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("G:G").Copy()
                    exlworksheet.Range("D1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("G1").EntireColumn.Delete()
                    exlworksheet.Range("G:G").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("B2:B3:C2:D2:E2:F2:G2:H2").Merge()
                    exlworksheet.Range("B2").Value = "المشتريات"
                    exlworksheet.Range("B2").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("B2").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("B2:B3:C2:C3:D2:D3:E2:E3:F2:F3:G2:G3:H2:H3").Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range("B2").Font.Size = 20
                    exlworksheet.Range("B2").Font.Bold = True
                    exlworksheet.Range("B2").Interior.Color = Color.WhiteSmoke
                    exlworksheet.Range("J2").Value = "=Count(G:G)"
                    Dim cn As Integer = exlworksheet.Range("J2").Value
                    With exlworksheet
                        .Range(.Cells(cn + 7, 7), .Cells(cn + 7, 8)).Merge()
                    End With
                    With exlworksheet
                        .Range(.Cells(cn + 8, 7), .Cells(cn + 8, 8)).Merge()
                    End With
                    exlworksheet.Cells(cn + 7, 6).Value = "=Sum(E6:E600)"
                    exlworksheet.Cells(cn + 7, 6).numberformat = "#        ريال"
                    exlworksheet.Cells(cn + 7, 6).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 7, 6).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 7, 7).value = "مجموع فواتير المشتريات :"
                    exlworksheet.Cells(cn + 8, 6).Value = "=Sum(D6:D600)"
                    exlworksheet.Cells(cn + 8, 6).numberformat = "0.0        ريال"
                    exlworksheet.Cells(cn + 8, 6).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 8, 6).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 8, 7).value = "مجموع ضرائب فواتير المشتريات :"
                    exlworksheet.Range("A4:A5:A6").EntireRow.Insert()
                    exlworksheet.Range("G5:F5").Merge()
                    exlworksheet.Range("F5").Value = ": من تاريخ"
                    exlworksheet.Range("F5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("F5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D5:E5").Merge()
                    exlworksheet.Range("D5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D5").Value = DateTimePicker1.Value
                    exlworksheet.Range("D5").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("G6:F6").Merge()
                    exlworksheet.Range("F6").Value = ": الى تاريخ"
                    exlworksheet.Range("F6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("F6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D6:E6").Merge()
                    exlworksheet.Range("D6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D6").Value = DateTimePicker2.Value.ToShortDateString
                    exlworksheet.Range("D6").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Cells(1, 2).columnwidth = 22
                    exlworksheet.Cells(1, 3).columnwidth = 25
                    exlworksheet.Cells(1, 4).columnwidth = 12
                    exlworksheet.Cells(1, 5).columnwidth = 12
                    exlworksheet.Cells(1, 6).columnwidth = 16
                    exlworksheet.Cells(1, 7).columnwidth = 15
                    exlworksheet.Cells(1, 8).columnwidth = 10
                    exlworksheet.Range("J2").Value = ""

                    'Sales
                    exlworksheet.Range("L1").EntireColumn.Delete()
                    exlworksheet.Range("N1").EntireColumn.Delete()
                    exlworksheet.Range("Q1").EntireColumn.Delete()
                    exlworksheet.Range("L:L").Copy()
                    exlworksheet.Range("Q1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("O:O").Copy()
                    exlworksheet.Range("L1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("P:P").Copy()
                    exlworksheet.Range("O1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("M:M").Copy()
                    exlworksheet.Range("P1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("M1").EntireColumn.Delete()
                    exlworksheet.Range("J1").EntireColumn.Delete()
                    exlworksheet.Range("N:N").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("K2:K3:L2:M2:N2:O2").Merge()
                    exlworksheet.Range("K2").Value = "المبيعات"
                    exlworksheet.Range("K2").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("K2").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("K2:K3:L2:L3:M2:M3:N2:N3:O2:O3").Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range("K2").Font.Size = 20
                    exlworksheet.Range("K2").Font.Bold = True
                    exlworksheet.Range("K2").Interior.Color = Color.LavenderBlush
                    exlworksheet.Range("R2").Value = "=Count(N:N)"
                    Dim cn1 As Integer = exlworksheet.Range("R2").Value
                    With exlworksheet
                        .Range(.Cells(cn1 + 10, 14), .Cells(cn1 + 10, 15)).Merge()
                    End With
                    With exlworksheet
                        .Range(.Cells(cn1 + 11, 14), .Cells(cn1 + 11, 15)).Merge()
                    End With
                    exlworksheet.Cells(cn1 + 10, 13).Value = "=Sum(L9:L600)"
                    exlworksheet.Cells(cn1 + 10, 13).numberformat = "#        ريال"
                    exlworksheet.Cells(cn1 + 10, 13).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 10, 13).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 10, 14).value = "مجموع فواتير المبيعات :"
                    exlworksheet.Cells(cn1 + 11, 13).Value = "=Sum(K9:K600)"
                    exlworksheet.Cells(cn1 + 11, 13).numberformat = "0.0        ريال"
                    exlworksheet.Cells(cn1 + 11, 13).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 11, 13).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 11, 14).value = "مجموع ضرائب فواتير المبيعات :"
                    exlworksheet.Range("N5:O5").Merge()
                    exlworksheet.Range("N5").Value = ": من تاريخ"
                    exlworksheet.Range("N5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("N5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M5:L5").Merge()
                    exlworksheet.Range("L5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("L5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("L5").Value = DateTimePicker1.Value.ToShortDateString
                    exlworksheet.Range("L5").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("N6:O6").Merge()
                    exlworksheet.Range("N6").Value = ": الى تاريخ"
                    exlworksheet.Range("N6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("N6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M6:L6").Merge()
                    exlworksheet.Range("L6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("L6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("L6").Value = DateTimePicker2.Value.ToShortDateString
                    exlworksheet.Range("L6").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("R2").Value = ""
                    exlworksheet.Cells(1, 11).columnwidth = 12
                    exlworksheet.Cells(1, 12).columnwidth = 12
                    exlworksheet.Cells(1, 13).columnwidth = 16
                    exlworksheet.Cells(1, 14).columnwidth = 15
                    exlworksheet.Cells(1, 15).columnwidth = 10
                    exlworksheet.Range(exlworksheet.Cells(8, 11), exlworksheet.Cells(cn1 + 8, 15)).Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range(exlworksheet.Cells(8, 2), exlworksheet.Cells(cn + 8, 8)).Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range("B:B").NumberFormat = "#"
                    exlworksheet.Cells(1, 1).Select()


                Catch
                    ' For Arabic Office
                    exlworksheet = exlworkbook.Sheets("ورقة1")

                    'Buys
                    RB_Form.DateTimePicker1.Value = DateTimePicker1.Value
                    RB_Form.DateTimePicker2.Value = DateTimePicker2.Value.AddDays(1)
                    RB_Form.Show()
                    RB_Form.Visible = False

                    'sales
                    RS_Form.DateTimePicker1.Value = DateTimePicker1.Value
                    RS_Form.DateTimePicker2.Value = DateTimePicker2.Value.AddDays(1)
                    RS_Form.Show()
                    RS_Form.Visible = False

                    'Buys
                    For colh = 0 To RB_Form.DataGridView1.ColumnCount - 1
                        exlworksheet.Cells(5, colh + 2) = RB_Form.DataGridView1.Columns(colh).HeaderText + ":"
                    Next

                    'Sales
                    For colh1 = 0 To RS_Form.DataGridView1.ColumnCount - 1
                        exlworksheet.Cells(5, colh1 + 16) = RS_Form.DataGridView1.Columns(colh1).HeaderText + ":"
                    Next

                    'Buys
                    For i As Integer = 0 To RB_Form.DataGridView1.RowCount - 1
                        For j As Integer = 0 To RB_Form.DataGridView1.ColumnCount - 1
                            exlworksheet.Cells(i + 6, j + 2) = RB_Form.DataGridView1.Rows(i).Cells(j).Value.ToString
                        Next
                    Next

                    'Sales
                    For i1 As Integer = 0 To RS_Form.DataGridView1.RowCount - 1
                        For j1 As Integer = 0 To RS_Form.DataGridView1.ColumnCount - 1
                            exlworksheet.Cells(i1 + 6, j1 + 16) = RS_Form.DataGridView1.Rows(i1).Cells(j1).Value.ToString
                        Next
                    Next

                    'Buys
                    exlworksheet.Range("B1").EntireColumn.Delete()
                    exlworksheet.Range("D1").EntireColumn.Delete()
                    exlworksheet.Range("G1").EntireColumn.Delete()
                    exlworksheet.Range("D1").EntireColumn.Insert()
                    exlworksheet.Range("G:G").Copy()
                    exlworksheet.Range("D1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("G1").EntireColumn.Delete()
                    exlworksheet.Range("C:C").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("B2:B3:C2:D2:E2:F2:G2:H2").Merge()
                    exlworksheet.Range("B2").Value = "المشتريات"
                    exlworksheet.Range("B2").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("B2").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("B2:B3:C2:C3:D2:D3:E2:E3:F2:F3:G2:G3:H2:H3").Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range("B2").Font.Size = 20
                    exlworksheet.Range("B2").Font.Bold = True
                    exlworksheet.Range("B2").Interior.Color = Color.WhiteSmoke
                    exlworksheet.Range("I2").Value = "=Count(C:C)"
                    Dim cn As Integer = exlworksheet.Range("I2").Value
                    With exlworksheet
                        .Range(.Cells(cn + 7, 2), .Cells(cn + 7, 3)).Merge()
                    End With
                    With exlworksheet
                        .Range(.Cells(cn + 8, 2), .Cells(cn + 8, 3)).Merge()
                    End With
                    exlworksheet.Cells(cn + 7, 4).Value = "=Sum(E6:E600)"
                    exlworksheet.Cells(cn + 7, 4).numberformat = "#        ريال"
                    exlworksheet.Cells(cn + 7, 4).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 7, 4).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 7, 2).value = "مجموع فواتير المشتريات :"
                    exlworksheet.Cells(cn + 8, 4).Value = "=SUM(F6:F600)"
                    exlworksheet.Cells(cn + 8, 4).numberformat = "0.0        ريال"
                    exlworksheet.Cells(cn + 8, 4).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 8, 4).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn + 8, 2).value = "مجموع ضرائب فواتير المشتريات :"
                    exlworksheet.Range("A4:A5:A6").EntireRow.Insert()
                    exlworksheet.Range("D5:E5").Merge()
                    exlworksheet.Range("D5").Value = ": من تاريخ"
                    exlworksheet.Range("D5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("G5:F5").Merge()
                    exlworksheet.Range("F5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("F5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("F5").Value = DateTimePicker1.Value
                    exlworksheet.Range("F5").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("D6:E6").Merge()
                    exlworksheet.Range("D6").Value = ": الى تاريخ"
                    exlworksheet.Range("D6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("D6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("G6:F6").Merge()
                    exlworksheet.Range("F6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("F6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("F6").Value = DateTimePicker2.Value.ToShortDateString
                    exlworksheet.Range("F6").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Cells(1, 2).columnwidth = 10
                    exlworksheet.Cells(1, 3).columnwidth = 15
                    exlworksheet.Cells(1, 4).columnwidth = 13
                    exlworksheet.Cells(1, 5).columnwidth = 12
                    exlworksheet.Cells(1, 6).columnwidth = 16
                    exlworksheet.Cells(1, 7).columnwidth = 20
                    exlworksheet.Cells(1, 8).columnwidth = 25
                    exlworksheet.Range("H:H").NumberFormat = "#"
                    exlworksheet.Range("I2").Value = ""

                    'Sales
                    exlworksheet.Range("M1").EntireColumn.Delete()
                    exlworksheet.Range("O1").EntireColumn.Delete()
                    exlworksheet.Range("R1").EntireColumn.Delete()
                    exlworksheet.Range("O1").EntireColumn.Insert()
                    exlworksheet.Range("R:R").Copy()
                    exlworksheet.Range("O1").PasteSpecial(XlPasteType.xlPasteAll)
                    exlworksheet.Range("R1").EntireColumn.Delete()
                    exlworksheet.Range("J1").EntireColumn.Delete()
                    exlworksheet.Range("J1").EntireColumn.Delete()
                    exlworksheet.Range("L:L").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("K2:K3:L2:M2:N2:O2").Merge()
                    exlworksheet.Range("K2").Value = "المبيعات"
                    exlworksheet.Range("K2").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("K2").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("K2:K3:L2:L3:M2:M3:N2:N3:O2:O3").Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range("K2").Font.Size = 20
                    exlworksheet.Range("K2").Font.Bold = True
                    exlworksheet.Range("K2").Interior.Color = Color.LavenderBlush
                    exlworksheet.Range("R2").Value = "=Count(L:L)"
                    Dim cn1 As Integer = exlworksheet.Range("R2").Value
                    With exlworksheet
                        .Range(.Cells(cn1 + 10, 11), .Cells(cn1 + 10, 12)).Merge()
                    End With
                    With exlworksheet
                        .Range(.Cells(cn1 + 11, 11), .Cells(cn1 + 11, 12)).Merge()
                    End With
                    exlworksheet.Cells(cn1 + 10, 13).Value = "=Sum(N9:N600)"
                    exlworksheet.Cells(cn1 + 10, 13).numberformat = "#        ريال"
                    exlworksheet.Cells(cn1 + 10, 13).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 10, 13).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 10, 11).value = "مجموع فواتير المبيعات :"
                    exlworksheet.Cells(cn1 + 11, 13).Value = "=Sum(O9:O600)"
                    exlworksheet.Cells(cn1 + 11, 13).numberformat = "0.0        ريال"
                    exlworksheet.Cells(cn1 + 11, 13).HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 11, 13).VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Cells(cn1 + 11, 11).value = "مجموع ضرائب فواتير المبيعات :"
                    exlworksheet.Range("K5:L5").Merge()
                    exlworksheet.Range("K5").Value = ": من تاريخ"
                    exlworksheet.Range("K5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("K5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M5:N5").Merge()
                    exlworksheet.Range("M5").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M5").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M5").Value = DateTimePicker1.Value.ToShortDateString
                    exlworksheet.Range("M5").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("K6:L6").Merge()
                    exlworksheet.Range("K6").Value = ": الى تاريخ"
                    exlworksheet.Range("K6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("K6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M6:N6").Merge()
                    exlworksheet.Range("M6").HorizontalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M6").VerticalAlignment = XlVAlign.xlVAlignCenter
                    exlworksheet.Range("M6").Value = DateTimePicker2.Value.ToShortDateString
                    exlworksheet.Range("M6").NumberFormat = "dd/MM/yyyy"
                    exlworksheet.Range("R2").Value = ""
                    exlworksheet.Cells(1, 11).columnwidth = 10
                    exlworksheet.Cells(1, 12).columnwidth = 15
                    exlworksheet.Cells(1, 13).columnwidth = 13
                    exlworksheet.Cells(1, 14).columnwidth = 12
                    exlworksheet.Cells(1, 15).columnwidth = 15
                    exlworksheet.Range(exlworksheet.Cells(8, 11), exlworksheet.Cells(cn1 + 8, 15)).Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Range(exlworksheet.Cells(8, 2), exlworksheet.Cells(cn + 8, 8)).Borders.Weight = BorderStyle.Fixed3D
                    exlworksheet.Cells(1, 1).Select()

                End Try



                SaveFileDialog1.Filter = "Excel Files|*.xlsx"
                SaveFileDialog1.FileName = "الضرائب"
                If SaveFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    Try
                        exlworksheet.SaveAs(SaveFileDialog1.FileName)
                        exlworkbook.Close()
                        exlapp.Quit()

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(exlapp)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(exlworkbook)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(exlworksheet)

                        Process.Start(SaveFileDialog1.FileName)
                    Catch
                        MsgBox("الرجاء التحقق ان الملف المحول مغلق",, "خطأ")
                    End Try
                Else

                End If
            Catch
                MsgBox("الرجاء اغلاق التقرير ومن ثم اعادة المحاولة مرة اخرى",, "لا يمكن التحويل")
            End Try
        Else
            MsgBox("!!! excel لايوجد حركات لتحويلها الى ملف", MessageBoxButtons.OK, "لا يمكن التحويل")
        End If
        exlapp = Nothing
        exlworkbook = Nothing
        exlworksheet = Nothing
        RB_Form.Close()
        RS_Form.Close()
    End Sub

    Sub Search2()
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
        Dim SearchDate1 As Date = DateTimePicker1.Value
        Dim SearchDate2 As Date = DateTimePicker2.Value.AddDays(1)
        dtTable2.Clear()
        da2 = New OleDbDataAdapter("select The_Table.ID, The_Table.Num, The_Table.Date1, The_Table.Describe1,
                                   The_Table.Price, The_Table.VAT, The_Table.Type, The_Table.Comer_Name , 
                                   Comers.Comer_Name1,Comers.Num_VAT
                                   from The_Table,Comers 
                                   where The_Table.Comer_Name = Comers.ID 
                                   and The_Table.Date1 > #" & SearchDate1.Year & "/" & SearchDate1.Month & "/" & SearchDate1.Day & "# and The_Table.Date1 < #" & SearchDate2.Year & "/" & SearchDate2.Month & "/" & SearchDate2.Day & "# order by Type", Con)

        da2.Fill(dtTable2)
        DataGridView1.DataSource = dtTable2
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "رقم الفاتورة"
        DataGridView1.Columns(2).HeaderText = "تاريخ الفاتورة"
        DataGridView1.Columns(3).HeaderText = "البيان"
        DataGridView1.Columns(4).HeaderText = "المبلغ"
        DataGridView1.Columns(5).HeaderText = "الضريبة"
        DataGridView1.Columns(6).HeaderText = "نوع الفاتورة"
        DataGridView1.Columns(2).DefaultCellStyle.Format = "dd/MM/yyyy"
        DataGridView1.Columns(5).DefaultCellStyle.Format = "0.#"
        DataGridView1.Columns(8).HeaderText = "اسم المورد"
        DataGridView1.Columns(9).HeaderText = "رقم السجل الضريبي"
        DataGridView1.Columns(3).Width = 250
        DataGridView1.Columns(8).Width = 150
        DataGridView1.Columns(6).Width = 144
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(9).Visible = False

        Label6.Text = (From row In
               DataGridView1.Rows.Cast(Of DataGridViewRow)() Where row.Cells(6).Value = "مبيعات" Select
                                                                               CDbl(row.Cells(5).Value)).Sum()
        Label9.Text = (From row In
                             DataGridView1.Rows.Cast(Of DataGridViewRow)() Where row.Cells(6).Value = "مشتريات" Select
                                                                               CDbl(row.Cells(5).Value)).Sum()
        Label10.Text = Convert.ToDouble(Label9.Text) * 2
        Label7.Text = Convert.ToDouble(Label9.Text) - Convert.ToDouble(Label10.Text)
        Label8.Text = Val(Label6.Text) + Val(Label7.Text)
        Label6.Text = Decimal.Parse(Label6.Text).ToString("0.#")
        Label7.Text = Decimal.Parse(Label7.Text).ToString("0.#")
        Label8.Text = Decimal.Parse(Label8.Text).ToString("0.#")
        If Label8.Text < 0 Then
            Label8.BackColor = Color.FromArgb(255, 128, 128)
            Label8.ForeColor = ForeColor.Black
        ElseIf Label8.Text > 0 Then
            Label8.BackColor = BackColor.Lime
            Label8.ForeColor = ForeColor.Black
        ElseIf Label8.Text = 0 Then
            Label8.BackColor = BackColor.White
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance        
        End If
    End Sub
    Private Sub RS_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Search2()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Search2()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        excel1()
    End Sub
End Class