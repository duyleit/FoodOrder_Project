Imports System
Imports System.Data
Imports Microsoft.Office.Interop
Imports System.Data.OleDb


Public Class FoodOrder_Frm

    Public status As Integer
    Private dtprint As DataTable
    ' Public tb As DataTable
    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        If Checkday() Then
            MessageBox.Show("Bạn Chỉ Được Đặt Hôm Nay Hoặc Không Được Quá 2 Ngày !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        rd_trua.Checked = True
        pnl_select.Show()
        'cb_dept.SelectedIndex = -1
        status = 0
        dgv_food.DataSource = ""
    End Sub

    Private Sub btn_pnl_select_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_pnl_select_close.Click
        pnl_select.Hide()
    End Sub

    Private Sub btn_upd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_upd.Click
        If Checkday() Then
            MessageBox.Show("Bạn Chỉ Được Cập Nhật Hôm Nay Hoặc Không Được Quá 2 Ngày !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        rd_trua.Checked = True
        pnl_select.Show()
        'cb_dept.SelectedIndex = -1
        btn_save.Enabled = False
        btn_cancel.Enabled = False
        lbl_pnl_select_title.Visible = False
        dgv_food.DataSource = ""
        status = 1
    End Sub
    Private Function getcon() As OleDbConnection
        Dim con As OleDbConnection
        con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.7.251\datcom$\FoodOrder.mdb")
        'con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|FoodOrder.mdb")
        Return con
    End Function
    Private Function Checkday() As Boolean  '20/3/2018
        Dim bflag As Boolean = False
        Dim timep As TimeSpan
        timep = dtp_date.Value.Date - Date.Now.Date
        'If timep.Days > 2 Then  'Cap nhat cho HC 2742018 bo sung so luong com con thieu
        ' If dtp_date.Value.Date.ToString("yyyy/MM/dd") < Date.Now.ToString("yyyy/MM/dd") Or timep.Days > 2 Then
        If dtp_date.Value.Date.ToString("yyyy/MM/dd") < Date.Now.ToString("yyyy/MM/dd") Then  'cap nhật ko gioi han ngày đặt cơm 19052018
            bflag = True
        End If
        Return bflag
    End Function
    Private Sub FoodOrder_Frm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '  Me.WindowState = FormWindowState.Maximized
        ' Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Dim tb As DataTable
        tb = New DataTable()

        ' Dim con As OleDbConnection
        ' con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\it\FoodOrder.mdb")
        'con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|FoodOrder.mdb")
        Dim con As OleDbConnection = getcon()
        Dim adapter As OleDbDataAdapter
        adapter = New OleDbDataAdapter("Select Ma,Ten From Dept Order by Ten", con)
        Try
            con.Open()
        Catch ex As Exception
            ' MessageBox.Show("Không Thể Kết Nối Dữ Liệu !!!")
            MessageBox.Show("Không Thể Kết Nối Dữ Liệu !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Dispose()
            Exit Sub
        End Try
        adapter.Fill(tb)
        dtp_date.Value = Date.Now.ToString("yyyy/MM/dd") '20/3/2018
        '  dtp_date.Value = Date.Now.ToString("MM") '3
        cb_dept.DisplayMember = "Ten"
        cb_dept.ValueMember = "Ma"
        cb_dept.DataSource = tb
        cb_dept.SelectedIndex = -1
        ' Cap nhat thuc don dong 28042018
        Try
            Dim img As Image = Image.FromFile("\\192.168.7.251\datcom$\menu.png")
            '   PictureBox1.Image = img
        Catch ex As Exception
            MessageBox.Show("Đang Cập Nhật Thực Đơn Mới !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Dispose()
            Exit Sub
        End Try

        btn_save.Enabled = False
        btn_cancel.Enabled = False
    End Sub

    Private Sub btn_pnl_select_ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_pnl_select_ok.Click

        Dim tb As DataTable
        tb = New DataTable()
        Dim dr As DataRow

        'Dim con As OleDbConnection
        ' con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\it\FoodOrder.mdb")

        Dim adapter As OleDbDataAdapter
        Dim asql As String
        asql = "select * from FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue
        If status = 0 Then 'Add
            txt_note.Enabled = True
            txt_note.Text = ""
            btn_save.Enabled = True
            btn_cancel.Enabled = True
            lbl_pnl_select_title.Visible = True
            If rd_trua.Checked Then
                'If CheckRecord(asql) Then
                '    MessageBox.Show("Dữ liệu này đã được thêm hôm nay. Bạn không thể thêm được nữa !!!")
                '    Exit Sub
                'End If
                lbl_pnl_select_title.Text = "Đặt Suất Ăn: Buổi Trưa"
                'adapter = New OleDbDataAdapter("Select Comman_trua,Phuman_trua,Comchay_trua,Chao_trua,Phuchay_trua From FoodQuantity where 1 > 2 ", con)
                'con.Open()
                'adapter.Fill(tb)
                tb.Columns.Add("comman_trua", GetType(Integer))
                tb.Columns.Add("phuman_trua", GetType(Integer))
                tb.Columns.Add("comchay_trua", GetType(Integer))
                tb.Columns.Add("chao_trua", GetType(Integer))
                tb.Columns.Add("phuchay_trua", GetType(Integer))
                dr = tb.NewRow()
                dr(0) = 0
                dr(1) = 0
                dr(2) = 0
                dr(3) = 0
                dr(4) = 0
                tb.Rows.Add(dr)
                dgv_food.DataSource = tb
                ' dgv_food.AllowUserToAddRows = True

                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                'For i As Integer = 0 To dgv_food.Columns.Count - 1
                '    dgv_food.Rows(0).Cells(i).Value = 0
                'Next
                dgv_food.Columns(0).HeaderText = "Cơm Mặn Trưa"
                dgv_food.Columns(1).HeaderText = "Phụ Mặn Trưa"
                dgv_food.Columns(2).HeaderText = "Cơm Chay Trưa"
                dgv_food.Columns(3).HeaderText = "Cháo Trưa"
                dgv_food.Columns(4).HeaderText = "Phụ Chay Trưa"
            ElseIf rd_chieu.Checked Then
                lbl_pnl_select_title.Text = "Đặt Suất Ăn: Buổi Chiều"
                ' adapter = New OleDbDataAdapter("Select Comman_chieu,Phuman_chieu,Phuchay_chieu From FoodQuantity where 1 > 2", con)
                ' con.Open()
                '  adapter.Fill(tb)
                tb.Columns.Add("comman_chieu", GetType(Integer))
                tb.Columns.Add("phuman_chieu", GetType(Integer))
                tb.Columns.Add("phuchay_chieu", GetType(Integer))
                dr = tb.NewRow()
                dr(0) = 0
                dr(1) = 0
                dr(2) = 0
                tb.Rows.Add(dr)
                dgv_food.DataSource = tb
                ' dgv_food.AllowUserToAddRows = True
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Cơm Mặn Chiều"
                dgv_food.Columns(1).HeaderText = "Phụ Mặn Chiều"
                dgv_food.Columns(2).HeaderText = "Phụ Chay Chiều"
            ElseIf rd_dem.Checked Then
                lbl_pnl_select_title.Text = "Đặt Suất Ăn: Buổi Đêm"
                ' adapter = New OleDbDataAdapter("Select Man_dem,Chay_dem From FoodQuantity where 1 > 2", con)
                'con.Open()
                ' adapter.Fill(tb)
                tb.Columns.Add("man_dem", GetType(Integer))
                tb.Columns.Add("chay_dem", GetType(Integer))
                dr = tb.NewRow()
                dr(0) = 0
                dr(1) = 0
                tb.Rows.Add(dr)
                dgv_food.DataSource = tb
                ' dgv_food.AllowUserToAddRows = True
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Mặn Suất Đêm"
                dgv_food.Columns(1).HeaderText = "Chay Suất Đêm"
            Else
                lbl_pnl_select_title.Text = "Đặt Suất Ăn: Cơm Khách"
                ' adapter = New OleDbDataAdapter("Select Donbanrieng,Buffet From FoodQuantity where 1 > 2", con)
                ' con.Open()
                ' adapter.Fill(tb)
                tb.Columns.Add("donbanrieng", GetType(Integer))
                tb.Columns.Add("buffet", GetType(Integer))
                dr = tb.NewRow()
                dr(0) = 0
                dr(1) = 0
                tb.Rows.Add(dr)
                dgv_food.DataSource = tb
                ' dgv_food.AllowUserToAddRows = True
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Dọn Bàn Riêng"
                dgv_food.Columns(1).HeaderText = "Ăn Buffet"
            End If
        Else  ' upd
            If cb_dept.SelectedIndex = -1 Then
                '  MessageBox.Show("Bạn Chưa Chọn Bộ Phận !!!")
                MessageBox.Show("Bạn Chưa Chọn Bộ Phận !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            asql = "select * from FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue
            If CheckExist(asql) <> True Then
                '  MessageBox.Show("Không Có Dữ Liệu Để Cập Nhật !!!")
                MessageBox.Show("Không Có Dữ Liệu Để Cập Nhật !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            ' Dim sqlstr As String
            'sqlstr = "Select Comman_trua,Phuman_trua,Comchay_trua,Chao_trua,Phuchay_trua From FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue
            txt_note.Enabled = True
            txt_note.Text = ""
            btn_save.Enabled = True
            btn_cancel.Enabled = True
            lbl_pnl_select_title.Visible = True
            If rd_trua.Checked Then

                lbl_pnl_select_title.Text = "Suất Ăn: Buổi Trưa"
                adapter = New OleDbDataAdapter("Select Comman_trua,Phuman_trua,Comchay_trua,Chao_trua,Phuchay_trua,ghichu From FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, getcon())
                getcon().Open()
                adapter.Fill(tb)

                dgv_food.DataSource = tb
                txt_note.Text = dgv_food.Rows(0).Cells(dgv_food.Columns.Count - 1).Value
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Cơm Mặn Trưa"
                dgv_food.Columns(1).HeaderText = "Phụ Mặn Trưa"
                dgv_food.Columns(2).HeaderText = "Cơm Chay Trưa"
                dgv_food.Columns(3).HeaderText = "Cháo Trưa"
                dgv_food.Columns(4).HeaderText = "Phụ Chay Trưa"
                dgv_food.Columns(5).Visible = False
            ElseIf rd_chieu.Checked Then
                lbl_pnl_select_title.Text = "Suất Ăn: Buổi Chiều"
                adapter = New OleDbDataAdapter("Select Comman_chieu,Phuman_chieu,Phuchay_chieu,ghichu From FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, getcon())
                getcon.Open()
                adapter.Fill(tb)

                dgv_food.DataSource = tb
                txt_note.Text = dgv_food.Rows(0).Cells(dgv_food.Columns.Count - 1).Value
                ' dgv_food.AllowUserToAddRows = True
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Cơm Mặn Chiều"
                dgv_food.Columns(1).HeaderText = "Phụ Mặn Chiều"
                dgv_food.Columns(2).HeaderText = "Phụ Chay Chiều"
                dgv_food.Columns(3).Visible = False
            ElseIf rd_dem.Checked Then
                lbl_pnl_select_title.Text = "Suất Ăn: Buổi Đêm"
                adapter = New OleDbDataAdapter("Select Man_dem,Chay_dem,ghichu From FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, getcon())
                getcon().Open()
                adapter.Fill(tb)

                dgv_food.DataSource = tb
                txt_note.Text = dgv_food.Rows(0).Cells(dgv_food.Columns.Count - 1).Value
                ' dgv_food.AllowUserToAddRows = True
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Mặn Suất Đêm"
                dgv_food.Columns(1).HeaderText = "Chay Suất Đêm"
                dgv_food.Columns(2).Visible = False
            Else
                lbl_pnl_select_title.Text = "Suất Ăn: Cơm Khách"
                adapter = New OleDbDataAdapter("Select Donbanrieng,Buffet,ghichu From FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, getcon())
                getcon.Open()
                adapter.Fill(tb)

                dgv_food.DataSource = tb
                txt_note.Text = dgv_food.Rows(0).Cells(dgv_food.Columns.Count - 1).Value
                ' dgv_food.AllowUserToAddRows = True
                For i As Integer = 0 To dgv_food.Columns.Count - 1
                    dgv_food.Columns(i).ReadOnly = False
                Next
                dgv_food.Columns(0).HeaderText = "Dọn Bàn Riêng"
                dgv_food.Columns(1).HeaderText = "Ăn Buffet"
                dgv_food.Columns(2).Visible = False
            End If

        End If
        pnl_select.Hide()
    End Sub

    Private Sub btn_excel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_excel.Click
        'Dim tb As DataTable
        'tb = New DataTable()

        'Dim con As OleDbConnection
        'con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\public\FoodOrder.mdb")

        'Dim adapter As OleDbDataAdapter
        'adapter = New OleDbDataAdapter("Select Ngaydat ,Trolydat ,Ten ,Comman_trua ,Phuman_trua ,Comchay_trua ,Chao_trua ,Phuchay_trua,comman_chieu,phuman_chieu,phuchay_chieu,man_dem,chay_dem,donbanrieng,buffet,ghichu From FoodQuantity,Dept Where FoodQuantity.Ma = Dept.Ma and ngaydat =#" + dtp_date.Text + "#", con)
        'con.Open()
        'adapter.Fill(tb)
        'cb_dept.SelectedIndex = -1
        pnl_select.Hide()
        Try
            If dtprint.Rows.Count = 0 Then

                MessageBox.Show("Chưa Có Dữ Liệu Để Xuất !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            print(dtprint)
        Catch ex As Exception

            MessageBox.Show("Chưa Có Dữ Liệu Để Xuất !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        cb_dept.SelectedIndex = -1
    End Sub
    Sub print(ByVal tb As DataTable)

        Dim xlCenter = -4108
        Dim xlLeft = -4131
        Dim xlRight = -4152
        Dim xlTop = -4160
        Dim xlBottom = -4107

        Dim ExcelApp As Excel.Application
        Dim ExcelWookBook As Excel.Workbook
        Dim ExcelWorkSheet As Excel.Worksheet
        Dim page As Double = 1

        Dim sheet_no As Integer = 0
        ' Dim xiao_ji As Integer = 0
        '  Dim zhong_ji As Integer = 0
        '  Dim dian_jia As Double = 0
        '   Dim zhong_dian_jia As Double = 0
        '   Dim start_row As Integer = 0
        '    Dim page As Double = 0

        ' variable define 
        Dim iBgnRow As Integer = 1
        Dim sCol As String = "Q"
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture

        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        ' setProcessBar(True, 1, ds_UPD04.Tables("UPD").Rows.Count, 0)
        '  setProcessBar(True, 1, dsPur92.Tables(0).Rows.Count, 0)

        Try
            ExcelApp = New Excel.Application()
            ExcelApp.Visible = False

            ExcelApp.UserControl = True
            ExcelApp.DisplayAlerts = False

            '   Dim data_grid_row As Integer = 0
            ' Dim hua_dan_bianhao As Integer = 1

            ExcelWookBook = ExcelApp.Workbooks.Add()
            ExcelWorkSheet = CType(ExcelWookBook.Sheets.Add(ExcelWookBook.Worksheets(sheet_no + 1), Type.Missing), Excel.Worksheet)
            ExcelWorkSheet.Name = page.ToString

            '    ExcelWorkSheet.PageSetup.TopMargin = 45
            '   ExcelWorkSheet.PageSetup.RightMargin = 8.53
            '   ExcelWorkSheet.PageSetup.LeftMargin = 9.2
            '   ExcelWorkSheet.PageSetup.BottomMargin = 10
            '  ExcelWorkSheet.PageSetup.HeaderMargin = 90


            ExcelApp.Columns(1).ColumnWidth = 4
            ExcelApp.Columns(2).ColumnWidth = 18
            ExcelApp.Columns(3).ColumnWidth = 15

            ExcelApp.Columns(4).ColumnWidth = 9
            ExcelApp.Columns(5).ColumnWidth = 9
            ExcelApp.Columns(6).ColumnWidth = 9
            ExcelApp.Columns(7).ColumnWidth = 9
            ExcelApp.Columns(8).ColumnWidth = 9

            ExcelApp.Columns(9).ColumnWidth = 9
            ExcelApp.Columns(10).ColumnWidth = 9
            ExcelApp.Columns(11).ColumnWidth = 9
            ExcelApp.Columns(12).ColumnWidth = 9

            ExcelApp.Columns(13).ColumnWidth = 5
            ExcelApp.Columns(14).ColumnWidth = 5

            ExcelApp.Columns(15).ColumnWidth = 13
            ExcelApp.Columns(16).ColumnWidth = 7

            ExcelApp.Columns(17).ColumnWidth = 30
            'title Columns
            ExcelApp.Cells(iBgnRow, 1) = "SỐ LƯỢNG SUẤT ĂN NGÀY " + dtp_date.Value.ToString("dd/MM/yyyy")
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Font.Bold = True
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Merge()
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Font.Size = 20
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).HorizontalAlignment = xlCenter

            iBgnRow += 1

            ExcelApp.Range("A2:A3").MergeCells = True
            ExcelApp.Range("A2:A3").VerticalAlignment = xlCenter
            ExcelApp.Range("B2:B3").MergeCells = True
            ExcelApp.Range("B2:B3").VerticalAlignment = xlCenter
            ExcelApp.Range("B2:B3").HorizontalAlignment = xlCenter
            ExcelApp.Range("C2:C3").MergeCells = True
            ExcelApp.Range("C2:C3").VerticalAlignment = xlCenter
            ExcelApp.Range("D2:I2").MergeCells = True
            ExcelApp.Range("D2:I3").HorizontalAlignment = xlCenter
            ExcelApp.Range("D2:I3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
            ExcelApp.Range("J2:L2").MergeCells = True
            ExcelApp.Range("J2:L3").HorizontalAlignment = xlCenter
            ExcelApp.Range("J2:L3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen)
            ExcelApp.Range("M2:N2").MergeCells = True
            ExcelApp.Range("M2:N3").HorizontalAlignment = xlCenter
            ExcelApp.Range("M2:N3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Coral)
            ExcelApp.Range("O2:P2").MergeCells = True
            ExcelApp.Range("O2:P3").HorizontalAlignment = xlCenter
            ExcelApp.Range("O2:P3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PaleVioletRed)
            ExcelApp.Range("Q2:Q3").MergeCells = True
            ExcelApp.Range("Q2:Q3").VerticalAlignment = xlCenter
            ExcelApp.Range("Q2:Q3").HorizontalAlignment = xlCenter


            ExcelApp.Cells(iBgnRow, 1) = "STT"
            ExcelApp.Cells(iBgnRow, 2) = "Bộ Phận"
            ExcelApp.Cells(iBgnRow, 3) = "Trợ Lý Báo Cơm"
            ExcelApp.Cells(iBgnRow, 4) = "Trưa"
            ExcelApp.Cells(iBgnRow + 1, 4) = "Cơm Mặn"
            ExcelApp.Cells(iBgnRow + 1, 5) = "Phụ Mặn"
            ExcelApp.Cells(iBgnRow + 1, 6) = "Cơm Chay"
            ExcelApp.Cells(iBgnRow + 1, 7) = "Cháo"
            ExcelApp.Cells(iBgnRow + 1, 8) = "Phụ Chay"
            ExcelApp.Cells(iBgnRow + 1, 9) = "Tổng Cộng"

            ExcelApp.Cells(iBgnRow, 10) = "Chiều"
            ExcelApp.Cells(iBgnRow + 1, 10) = "Cơm Mặn"
            ExcelApp.Cells(iBgnRow + 1, 11) = "Phụ Mặn"
            ExcelApp.Cells(iBgnRow + 1, 12) = "Phụ Chay"
            ExcelApp.Cells(iBgnRow, 13) = "Đêm"
            ExcelApp.Cells(iBgnRow + 1, 13) = "Mặn"
            ExcelApp.Cells(iBgnRow + 1, 14) = "Chay"
            ExcelApp.Cells(iBgnRow, 15) = "Cơm Khách"
            ExcelApp.Cells(iBgnRow + 1, 15) = "Dọn Bàn Riêng"
            ExcelApp.Cells(iBgnRow + 1, 16) = "Buffet"
            ExcelApp.Cells(iBgnRow, 17) = "Ghi Chú"


            ' ExcelApp.Range("A1:" + sCol + iBgnRow.ToString).VerticalAlignment = Excel.XlHAlign.xlHAlignCenter


            iBgnRow += 2
            Dim stt As Integer = 1
            For rRow As Integer = 0 To tb.Rows.Count - 1

                ExcelApp.Cells(iBgnRow, 1) = stt
                ExcelApp.Cells(iBgnRow, 2) = tb.Rows(rRow).Item("ten").ToString()
                ExcelApp.Cells(iBgnRow, 3) = tb.Rows(rRow).Item("trolydat").ToString()

                ExcelApp.Cells(iBgnRow, 4) = tb.Rows(rRow).Item("comman_trua").ToString()
                ExcelApp.Cells(iBgnRow, 5) = tb.Rows(rRow).Item("phuman_trua").ToString()
                ExcelApp.Cells(iBgnRow, 6) = tb.Rows(rRow).Item("comchay_trua").ToString()
                ExcelApp.Cells(iBgnRow, 7) = tb.Rows(rRow).Item("chao_trua").ToString()
                ExcelApp.Cells(iBgnRow, 8) = tb.Rows(rRow).Item("phuchay_trua").ToString()

                ExcelApp.Cells(iBgnRow, 9) = tb.Rows(rRow).Item("comman_trua") + tb.Rows(rRow).Item("phuman_trua") + tb.Rows(rRow).Item("comchay_trua") + tb.Rows(rRow).Item("chao_trua") + tb.Rows(rRow).Item("phuchay_trua")

                ExcelApp.Cells(iBgnRow, 10) = tb.Rows(rRow).Item("comman_chieu").ToString()
                ExcelApp.Cells(iBgnRow, 11) = tb.Rows(rRow).Item("phuman_chieu").ToString()
                ExcelApp.Cells(iBgnRow, 12) = tb.Rows(rRow).Item("phuchay_chieu").ToString()

                ExcelApp.Cells(iBgnRow, 13) = tb.Rows(rRow).Item("man_dem").ToString()
                ExcelApp.Cells(iBgnRow, 14) = tb.Rows(rRow).Item("chay_dem").ToString()

                ExcelApp.Cells(iBgnRow, 15) = tb.Rows(rRow).Item("donbanrieng").ToString()
                ExcelApp.Cells(iBgnRow, 16) = tb.Rows(rRow).Item("buffet").ToString()

                ExcelApp.Cells(iBgnRow, 17) = tb.Rows(rRow).Item("ghichu").ToString()
                stt += 1
                iBgnRow += 1
            Next
            ExcelApp.Range("I2:I" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("Q2:Q" & iBgnRow - 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown)
            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).MergeCells = True
            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Bold = True
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("D4:P" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Cells(iBgnRow, 1) = "Tổng Cột"
            ExcelApp.Cells(iBgnRow, 4) = "=SUM(D4:D" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 5) = "=SUM(E4:E" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 6) = "=SUM(F4:F" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 7) = "=SUM(G4:G" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 8) = "=SUM(H4:H" + (iBgnRow - 1).ToString + ")"


            ExcelApp.Cells(iBgnRow, 10) = "=SUM(J4:J" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 11) = "=SUM(K4:K" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 12) = "=SUM(L4:L" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 13) = "=SUM(M4:M" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 14) = "=SUM(N4:N" + (iBgnRow - 1).ToString + ")"
            iBgnRow += 1

            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).MergeCells = True
            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Bold = True
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Color = Color.Red

            ' ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Cells(iBgnRow, 1) = "Tổng Chung"
            ExcelApp.Range("D" & iBgnRow & ":H" & iBgnRow).MergeCells = True
            ExcelApp.Range("D" & iBgnRow & ":H" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Range("J" & iBgnRow & ":L" & iBgnRow).MergeCells = True
            ExcelApp.Range("J" & iBgnRow & ":L" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Range("M" & iBgnRow & ":N" & iBgnRow).MergeCells = True
            ExcelApp.Range("M" & iBgnRow & ":N" & iBgnRow).HorizontalAlignment = xlCenter

            ExcelApp.Cells(iBgnRow, 4) = "=SUM(D" + (iBgnRow - 1).ToString + ":H" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 10) = "=SUM(J" + (iBgnRow - 1).ToString + ":L" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 13) = "=SUM(M" + (iBgnRow - 1).ToString + ":N" + (iBgnRow - 1).ToString + ")"

            iBgnRow += 1
            ExcelApp.Range("A2:" & sCol & (iBgnRow - 1).ToString).Borders.LineStyle = 1 'Excel.XlLineStyle.xlContinuous
            ExcelApp.Range("D5:E" + iBgnRow.ToString).WrapText = True
            Dim icount As Integer
            icount = CountDept() '20/3/2018
            If tb.Rows.Count < icount And cb_dept.SelectedIndex = -1 Then
                Dim Sqlstr As String
                Dim arr(icount) As String
                Sqlstr = "select ten from dept where ma not in(select distinct dept.ma  from dept,foodquantity where dept.ma=foodquantity.ma and ngaydat =#" & dtp_date.Text & "#)"
                iBgnRow += 2
                ExcelApp.Cells(iBgnRow, 2) = "Những Bộ Phận Còn Chưa Đặt Cơm: "
                ExcelApp.Range("B" & iBgnRow & ":C" & iBgnRow).Font.Bold = True
                ExcelApp.Range("B" & iBgnRow & ":C" & iBgnRow).Font.Color = Color.Red
                iBgnRow += 1
                arr = Split(GetMissDept(Sqlstr), ",", icount - tb.Rows.Count)
                For i As Integer = 0 To icount - tb.Rows.Count - 1
                    ExcelApp.Cells(iBgnRow, 2) = i + 1 & " " & arr(i)
                    iBgnRow += 1
                Next
            End If
            ExcelApp.Range("A4:" + sCol + iBgnRow.ToString).Font.Size = 11
            ExcelApp.Range("A1:" + sCol + iBgnRow.ToString).Font.Name = "Times New Roman"

            ExcelApp.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function GetMissDept(ByVal Sqlstr As String) As String
        Dim strdepts As String = ""

        'Dim con As OleDbConnection
        'con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\it\FoodOrder.mdb")
        Dim con As OleDbConnection = getcon()
        Dim SqlCmd As OleDbCommand
        SqlCmd = New OleDbCommand(Sqlstr, con)
        con.Open()
        Dim Dtr1 As OleDb.OleDbDataReader = SqlCmd.ExecuteReader()
        If Not Dtr1 Is Nothing Then
            While Dtr1.Read()
                strdepts &= Dtr1.GetValue(0) & " , "
            End While
            Dtr1.Close()
        End If
        Return strdepts
    End Function


    Private Sub btn_query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_query.Click
        'pnl_select.Show()
        'status = 2
        '   Dim tb As DataTable
        pnl_select.Hide()
        btn_save.Enabled = False
        btn_cancel.Enabled = False
        txt_note.Enabled = False
        'cb_dept.SelectedIndex = -1
        txt_note.Text = "Bộ phận nào không ăn cơm vui lòng nhập số 0 vào đúng vị trí của bộ phần mình và ghi chú cắt cơm ở dòng ghi chú.Đối với các bộ phận không báo cơm đúng thời gian quy định nhà ăn sẽ không phục vụ phần cơm của bộ phận đó."
        lbl_pnl_select_title.Visible = False

        dtprint = New DataTable()
        dgv_food.DataSource = ""
        dgv_food.AllowUserToAddRows = False
        ' Dim con As OleDbConnection
        ' con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\it\FoodOrder.mdb")
        Dim con As OleDbConnection = getcon()
        Dim adapter As OleDbDataAdapter
        'Dim ss As String = "Select Ngaydat,Trolydat,Ten,Comman_trua ,Phuman_trua ,Comchay_trua ,Chao_trua ,Phuchay_trua,comman_chieu,phuman_chieu,phuchay_chieu,man_dem,chay_dem,donbanrieng,buffet,ghichu From FoodQuantity,Dept Where FoodQuantity.Ma = Dept.Ma and ngaydat =#" + dtp_date.Text + "#"
        If cb_dept.SelectedIndex = -1 Then
            adapter = New OleDbDataAdapter("Select Ngaydat,Trolydat,Ten,Comman_trua ,Phuman_trua ,Comchay_trua ,Chao_trua ,Phuchay_trua,comman_chieu,phuman_chieu,phuchay_chieu,man_dem,chay_dem,donbanrieng,buffet,ghichu From FoodQuantity,Dept Where FoodQuantity.Ma = Dept.Ma and ngaydat =#" + dtp_date.Text + "#", con)
        Else
            adapter = New OleDbDataAdapter("Select Ngaydat,Trolydat,Ten,Comman_trua ,Phuman_trua ,Comchay_trua ,Chao_trua ,Phuchay_trua,comman_chieu,phuman_chieu,phuchay_chieu,man_dem,chay_dem,donbanrieng,buffet,ghichu From FoodQuantity,Dept Where FoodQuantity.Ma = Dept.Ma and ngaydat =#" + dtp_date.Text + "# and Dept.Ma =" & cb_dept.SelectedValue, con)
        End If
        con.Open()
        adapter.Fill(dtprint)
        dgv_food.DataSource = dtprint
        For i As Integer = 0 To dgv_food.Columns.Count - 1
            dgv_food.Columns(i).ReadOnly = True
        Next
        dgv_food.Columns(0).HeaderText = "Ngày Đặt"
        dgv_food.Columns(1).HeaderText = "Người Đặt"
        dgv_food.Columns(2).HeaderText = "Bộ Phận"

        dgv_food.Columns(3).HeaderText = "Cơm Mặn Trưa"
        dgv_food.Columns(4).HeaderText = "Phụ Mặn Trưa"
        dgv_food.Columns(5).HeaderText = "Cơm Chay Trưa"
        dgv_food.Columns(6).HeaderText = "Cháo Trưa"
        dgv_food.Columns(7).HeaderText = "Phụ Chay Trưa"

        dgv_food.Columns(8).HeaderText = "Cơm Mặn Chiều"
        dgv_food.Columns(9).HeaderText = "Phụ Mặn Chiều"
        dgv_food.Columns(10).HeaderText = "Phụ Chay Chiều"

        dgv_food.Columns(11).HeaderText = "Mặn Suất Đêm"
        dgv_food.Columns(12).HeaderText = "Chay Suất Đêm"

        dgv_food.Columns(13).HeaderText = "Dọn Bàn Riêng"
        dgv_food.Columns(14).HeaderText = "Ăn Buffet"
        dgv_food.Columns(15).HeaderText = "Ghi Chú"

        If dtprint.Rows.Count = 0 Then
            '  MessageBox.Show("Không Có Dữ Liệu !!!")
            MessageBox.Show("Không Có Dữ Liệu !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dgv_food.DataSource = ""
            Exit Sub
        End If
        'cb_dept.SelectedIndex = -1
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click

        'Dim con As OleDbConnection
        ' con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\it\FoodOrder.mdb")
        Dim cmd As OleDbCommand
        Dim con As OleDbConnection = getcon()
        Dim arrint(10) As Integer
        Dim ngay As Date
        Dim asql As String

        If Checkday() Then
            MessageBox.Show("Bạn Chỉ Được Đặt Hôm Nay Hoặc Không Được Quá 2 Ngày !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If cb_dept.SelectedIndex = -1 Then
            '  MessageBox.Show("Bạn Chưa Chọn Bộ Phận Để Đặt Cơm !!!")
            MessageBox.Show("Bạn Chưa Chọn Bộ Phận Để Đặt Cơm !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsDBNull(dgv_food) Then
            'MessageBox.Show("Không Có Dữ Liệu Để Lưu !!!")
            MessageBox.Show("Không Có Dữ Liệu Để Lưu !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        asql = "select * from FoodQuantity where ngaydat =#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue

        If MessageBox.Show("Bạn Có Muốn Lưu Dữ Liệu Vừa Nhập ?", "Chương Trình Đặt Cơm", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) <> MsgBoxResult.Ok Then
            Exit Sub
        End If
        'If dgv_food.Rows.Count < 1 Then
        '    Exit Sub
        'End If
        If status = 0 Then
            If rd_trua.Checked Then
                For i As Integer = 1 To dgv_food.Columns.Count
                    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                        arrint(i) = 0
                    Else
                        arrint(i) = dgv_food(i - 1, 0).Value
                    End If
                Next
                ngay = dtp_date.Text
                arrint(0) = cb_dept.SelectedValue
                If CheckExist(asql) Then
                    Dim sqlstr As String
                    sqlstr = "update FoodQuantity Set trolydat=@trolydat,comman_trua=@comman_trua,phuman_trua=@phuman_trua,comchay_trua=@comchay_trua,chao_trua=@chao_trua,phuchay_trua=@phuchay_trua,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue
                    cmd = New OleDbCommand(sqlstr, con)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@comman_trua", arrint(1))
                    cmd.Parameters.AddWithValue("@phuman_trua", arrint(2))
                    cmd.Parameters.AddWithValue("@comchay_trua", arrint(3))
                    cmd.Parameters.AddWithValue("@chao_trua", arrint(4))
                    cmd.Parameters.AddWithValue("@phuchay_trua", arrint(5))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    '  MessageBox.Show("Lưu Thành Công !!!")
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else 'them moi ma chua co row 
                    cmd = New OleDbCommand("insert into FoodQuantity(ngaydat,trolydat,ma,comman_trua,phuman_trua,comchay_trua,chao_trua,phuchay_trua,ghichu) values (@ngaydat,@trolydat,@ma,@comman_trua,@phuman_trua,@comchay_trua,@chao_trua,@phuchay_trua,@ghichu)", con)
                    cmd.Parameters.AddWithValue("@ngaydat", ngay)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName.ToString())
                    cmd.Parameters.AddWithValue("@ma", arrint(0))
                    cmd.Parameters.AddWithValue("@comman_trua", arrint(1))
                    cmd.Parameters.AddWithValue("@phuman_trua", arrint(2))
                    cmd.Parameters.AddWithValue("@comchay_trua", arrint(3))
                    cmd.Parameters.AddWithValue("@chao_trua", arrint(4))
                    cmd.Parameters.AddWithValue("@phuchay_trua", arrint(5))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            ElseIf rd_chieu.Checked Then
                For i As Integer = 1 To dgv_food.Columns.Count
                    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                        arrint(i) = 0
                    Else
                        arrint(i) = dgv_food(i - 1, 0).Value
                    End If
                Next
                ngay = dtp_date.Text
                arrint(0) = cb_dept.SelectedValue

                If CheckExist(asql) Then
                    cmd = New OleDbCommand("Update FoodQuantity Set trolydat=@trolydat,comman_chieu=@comman_chieu,phuman_chieu=@phuman_chieu,phuchay_chieu=@phuchay_chieu,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, con)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@comman_chieu", arrint(1))
                    cmd.Parameters.AddWithValue("@phuman_chieu", arrint(2))
                    cmd.Parameters.AddWithValue("@phuchay_chieu", arrint(3))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else 'them moi ma chua co row 
                    cmd = New OleDbCommand("insert into FoodQuantity(ngaydat,trolydat,ma,comman_chieu,phuman_chieu,phuchay_chieu,ghichu) values (@ngaydat,@trolydat,@ma,@comman_chieu,@phuman_chieu,@phuchay_chieu,@ghichu)", con)
                    cmd.Parameters.AddWithValue("@ngaydat", ngay)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@ma", arrint(0))
                    cmd.Parameters.AddWithValue("@comman_chieu", arrint(1))
                    cmd.Parameters.AddWithValue("@phuman_chieu", arrint(2))
                    cmd.Parameters.AddWithValue("@phuchay_chieu", arrint(3))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            ElseIf rd_dem.Checked Then
                For i As Integer = 1 To dgv_food.Columns.Count
                    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                        arrint(i) = 0
                    Else
                        arrint(i) = dgv_food(i - 1, 0).Value
                    End If
                Next
                ngay = dtp_date.Text
                arrint(0) = cb_dept.SelectedValue

                If CheckExist(asql) Then
                    cmd = New OleDbCommand("update FoodQuantity set trolydat=@trolydat,man_dem=@man_dem, chay_dem=@chay_dem,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, con)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@man_dem", arrint(1))
                    cmd.Parameters.AddWithValue("@chay_dem", arrint(2))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else 'them moi ma chua co row 
                    cmd = New OleDbCommand("insert into FoodQuantity(ngaydat,trolydat,ma,man_dem,chay_dem,ghichu) values (@ngaydat,@trolydat,@ma,@man_dem,@chay_dem,@ghichu)", con)
                    cmd.Parameters.AddWithValue("@ngaydat", ngay)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@ma", arrint(0))
                    cmd.Parameters.AddWithValue("@man_dem", arrint(1))
                    cmd.Parameters.AddWithValue("@chay_dem", arrint(2))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)
                    If con.State = ConnectionState.Closed Then
                        con.Open()
                    End If
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else 'com khach
                For i As Integer = 1 To dgv_food.Columns.Count
                    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                        arrint(i) = 0
                    Else
                        arrint(i) = dgv_food(i - 1, 0).Value
                    End If
                Next
                ngay = dtp_date.Text
                arrint(0) = cb_dept.SelectedValue

                If CheckExist(asql) Then
                    cmd = New OleDbCommand("update FoodQuantity set trolydat=@trolydat,donbanrieng=@donbanrieng,buffet=@buffet,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, con)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@donbanrieng", arrint(1))
                    cmd.Parameters.AddWithValue("@buffet", arrint(2))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else 'them moi ma chua co row 
                    cmd = New OleDbCommand("insert into FoodQuantity(ngaydat,trolydat,ma,donbanrieng,buffet,ghichu) values (@ngaydat,@trolydat,@ma,@donbanrieng,@buffet,@ghichu)", con)
                    cmd.Parameters.AddWithValue("@ngaydat", ngay)
                    cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                    cmd.Parameters.AddWithValue("@ma", arrint(0))
                    cmd.Parameters.AddWithValue("@donbanrieng", arrint(1))
                    cmd.Parameters.AddWithValue("@buffet", arrint(2))
                    cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            End If
        Else 'luu upd
            If rd_trua.Checked Then
                'For i As Integer = 1 To dgv_food.Columns.Count
                '    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                '        arrint(i) = 0
                '    Else
                '        arrint(i) = dgv_food(i - 1, 0).Value
                '    End If
                'Next
                'ngay = dtp_date.Text
                'arrint(0) = cb_dept.SelectedValue

                Dim sqlstr As String
                sqlstr = "update FoodQuantity Set trolydat=@trolydat,comman_trua=@comman_trua,phuman_trua=@phuman_trua,comchay_trua=@comchay_trua,chao_trua=@chao_trua,phuchay_trua=@phuchay_trua,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue
                cmd = New OleDbCommand(sqlstr, con)
                cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                cmd.Parameters.AddWithValue("@comman_trua", dgv_food.Rows(0).Cells(0).Value)
                cmd.Parameters.AddWithValue("@phuman_trua", dgv_food.Rows(0).Cells(1).Value)
                cmd.Parameters.AddWithValue("@comchay_trua", dgv_food.Rows(0).Cells(2).Value)
                cmd.Parameters.AddWithValue("@chao_trua", dgv_food.Rows(0).Cells(3).Value)
                cmd.Parameters.AddWithValue("@phuchay_trua", dgv_food.Rows(0).Cells(4).Value)
                cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                con.Open()
                cmd.ExecuteNonQuery()
                'MessageBox.Show("Cập Nhật Thành Công !!!")
                MessageBox.Show("Cập Nhật Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf rd_chieu.Checked Then
                'For i As Integer = 1 To dgv_food.Columns.Count
                '    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                '        arrint(i) = 0
                '    Else
                '        arrint(i) = dgv_food(i - 1, 0).Value
                '    End If
                'Next
                'ngay = dtp_date.Text
                'arrint(0) = cb_dept.SelectedValue
                cmd = New OleDbCommand("Update FoodQuantity Set trolydat=@trolydat,comman_chieu=@comman_chieu,phuman_chieu=@phuman_chieu,phuchay_chieu=@phuchay_chieu,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, con)
                cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                cmd.Parameters.AddWithValue("@comman_chieu", dgv_food.Rows(0).Cells(0).Value)
                cmd.Parameters.AddWithValue("@phuman_chieu", dgv_food.Rows(0).Cells(1).Value)
                cmd.Parameters.AddWithValue("@phuchay_chieu", dgv_food.Rows(0).Cells(2).Value)
                cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                con.Open()
                cmd.ExecuteNonQuery()
                MessageBox.Show("Cập Nhật Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ElseIf rd_dem.Checked Then
                'For i As Integer = 1 To dgv_food.Columns.Count
                '    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                '        arrint(i) = 0
                '    Else
                '        arrint(i) = dgv_food(i - 1, 0).Value
                '    End If
                'Next
                'ngay = dtp_date.Text
                'arrint(0) = cb_dept.SelectedValue
                cmd = New OleDbCommand("update FoodQuantity set trolydat=@trolydat,man_dem=@man_dem, chay_dem=@chay_dem,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, con)
                cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                cmd.Parameters.AddWithValue("@man_dem", dgv_food.Rows(0).Cells(0).Value)
                cmd.Parameters.AddWithValue("@chay_dem", dgv_food.Rows(0).Cells(1).Value)
                cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                con.Open()
                cmd.ExecuteNonQuery()
                MessageBox.Show("Cập Nhật Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else 'com khach
                'For i As Integer = 1 To dgv_food.Columns.Count
                '    If IsDBNull(dgv_food(i - 1, 0).Value) Then
                '        arrint(i) = 0
                '    Else
                '        arrint(i) = dgv_food(i - 1, 0).Value
                '    End If
                'Next
                'ngay = dtp_date.Text
                'arrint(0) = cb_dept.SelectedValue
                cmd = New OleDbCommand("update FoodQuantity set trolydat=@trolydat,donbanrieng=@donbanrieng,buffet=@buffet,ghichu=@ghichu where ngaydat=#" & dtp_date.Text & "# and ma=" & cb_dept.SelectedValue, con)
                cmd.Parameters.AddWithValue("@trolydat", Environment.UserName)
                cmd.Parameters.AddWithValue("@donbanrieng", dgv_food.Rows(0).Cells(0).Value)
                cmd.Parameters.AddWithValue("@buffet", dgv_food.Rows(0).Cells(1).Value)
                cmd.Parameters.AddWithValue("@ghichu", txt_note.Text)

                con.Open()
                cmd.ExecuteNonQuery()
                MessageBox.Show("Cập Nhật Thành Công !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
        dgv_food.DataSource = ""
        txt_note.Text = ""
        txt_note.Enabled = False
        lbl_pnl_select_title.Visible = False
        btn_save.Enabled = False
        btn_cancel.Enabled = False
        cb_dept.SelectedIndex = -1
    End Sub

    Private Sub dgv_food_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgv_food.DataError
        'MessageBox.Show("Bạn Chỉ Được Nhập Số !!!")
        MessageBox.Show("Bạn Chỉ Được Nhập Số !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub
    Private Function CheckExist(ByVal ASql As String) As Boolean

        Dim isExists As Boolean = False

        'Dim con As OleDbConnection
        'con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.12\it\FoodOrder.mdb")
        Dim con As OleDbConnection = getcon()
        Dim SqlCmd As OleDbCommand
        SqlCmd = New OleDbCommand(ASql, con)
        con.Open()
        Dim Dtr1 As OleDb.OleDbDataReader = SqlCmd.ExecuteReader()
        If Dtr1.HasRows Then
            isExists = True
        Else
            isExists = False
        End If
        Dtr1.Close()
        Return isExists
    End Function
    Private Function CountDept() As Integer
        Dim icount As Integer
        Dim con As OleDbConnection = getcon()
        Dim SqlCmd As OleDbCommand
        SqlCmd = New OleDbCommand("Select count(Ma) From Dept", con)
        con.Open()
        icount = SqlCmd.ExecuteScalar()
        Return icount
    End Function

    'Private Function CheckRecord(ByVal ASql As String) As Boolean
    '    Dim dr1 As SqlClient.SqlDataReader = getDataReader_38(ASql, conn)
    '    Dim isExists As Boolean = False

    '    If Not dr1 Is Nothing Then
    '        While dr1.Read()
    '            If Convert.IsDBNull(dr1.GetValue(0)) Then 'ctype
    '                isExists = True
    '            End If
    '            Exit While
    '        End While
    '        dr1.Close()
    '    End If
    '    Return isExists
    'End Function

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        If dgv_food.Rows.Count > 0 Then
            dgv_food.DataSource = ""
        End If
        txt_note.Text = ""
        txt_note.Enabled = False
        btn_save.Enabled = False
        cb_dept.SelectedIndex = -1
    End Sub

    Private Sub btn_ex2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ex2.Click

        pnl_select.Hide()
        Try
            If dtprint.Rows.Count = 0 Then
                MessageBox.Show("Chưa Có Dữ Liệu Để Xuất !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            print1(dtprint)
        Catch ex As Exception
            MessageBox.Show("Chưa Có Dữ Liệu Để Xuất !!!", "Chương Trình Đặt Cơm", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        cb_dept.SelectedIndex = -1
    End Sub

    Sub print1(ByVal tb As DataTable)

        'Ma ngoai Tong Van Phong: 29->37
        'Dim HT(100) As String '/// Mảng chứa được 101 phần tử từ 0 đến 100
        'Dim Diem(0 To 100) As Single '/// Mảng chứa được 101 phần tử từ 0 đến 100
        'Dim MaTran1(4, 4) As Single '/// Ma trận (mảng 2 chiều) có 5 hàng 5 cột
        'Dim MaTran2(0 To 4, 0 To 4) As Single '/// Mảng 2 chiều có 5 hàng, 5 cột

        Dim xlCenter = -4108
        Dim xlLeft = -4131
        Dim xlRight = -4152
        Dim xlTop = -4160
        Dim xlBottom = -4107

        Dim ExcelApp As Excel.Application
        Dim ExcelWookBook As Excel.Workbook
        Dim ExcelWorkSheet As Excel.Worksheet

        Dim sheet_no As Integer = 0

        Dim iBgnRow As Integer = 1
        Dim sCol As String = "Q"
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture

        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try
            ExcelApp = New Excel.Application()
            ExcelApp.Visible = False

            ExcelApp.UserControl = True
            ExcelApp.DisplayAlerts = False



            ExcelWookBook = ExcelApp.Workbooks.Add()
            ExcelWorkSheet = CType(ExcelWookBook.Sheets.Add(ExcelWookBook.Worksheets(sheet_no + 1), Type.Missing), Excel.Worksheet)

            ExcelApp.Columns(1).ColumnWidth = 4
            ExcelApp.Columns(2).ColumnWidth = 18
            ExcelApp.Columns(3).ColumnWidth = 15

            ExcelApp.Columns(4).ColumnWidth = 9
            ExcelApp.Columns(5).ColumnWidth = 9
            ExcelApp.Columns(6).ColumnWidth = 9
            ExcelApp.Columns(7).ColumnWidth = 9
            ExcelApp.Columns(8).ColumnWidth = 9

            ExcelApp.Columns(9).ColumnWidth = 9
            ExcelApp.Columns(10).ColumnWidth = 9
            ExcelApp.Columns(11).ColumnWidth = 9
            ExcelApp.Columns(12).ColumnWidth = 9

            ExcelApp.Columns(13).ColumnWidth = 5
            ExcelApp.Columns(14).ColumnWidth = 5

            ExcelApp.Columns(15).ColumnWidth = 13
            ExcelApp.Columns(16).ColumnWidth = 7

            ExcelApp.Columns(17).ColumnWidth = 30
            'title Columns
            ExcelApp.Cells(iBgnRow, 1) = "SỐ LƯỢNG SUẤT ĂN NGÀY " + dtp_date.Value.ToString("dd/MM/yyyy")
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Font.Bold = True
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Merge()
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).Font.Size = 20
            ExcelApp.Range("A" + iBgnRow.ToString + ":" & sCol & iBgnRow.ToString).HorizontalAlignment = xlCenter

            iBgnRow += 1

            ExcelApp.Range("A2:A3").MergeCells = True
            ExcelApp.Range("A2:A3").VerticalAlignment = xlCenter
            ExcelApp.Range("B2:B3").MergeCells = True
            ExcelApp.Range("B2:B3").VerticalAlignment = xlCenter
            ExcelApp.Range("B2:B3").HorizontalAlignment = xlCenter
            ExcelApp.Range("C2:C3").MergeCells = True
            ExcelApp.Range("C2:C3").VerticalAlignment = xlCenter
            ExcelApp.Range("D2:I2").MergeCells = True
            ExcelApp.Range("D2:I3").HorizontalAlignment = xlCenter
            ExcelApp.Range("D2:I3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
            ExcelApp.Range("J2:L2").MergeCells = True
            ExcelApp.Range("J2:L3").HorizontalAlignment = xlCenter
            ExcelApp.Range("J2:L3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen)
            ExcelApp.Range("M2:N2").MergeCells = True
            ExcelApp.Range("M2:N3").HorizontalAlignment = xlCenter
            ExcelApp.Range("M2:N3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Coral)
            ExcelApp.Range("O2:P2").MergeCells = True
            ExcelApp.Range("O2:P3").HorizontalAlignment = xlCenter
            ExcelApp.Range("O2:P3").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PaleVioletRed)
            ExcelApp.Range("Q2:Q3").MergeCells = True
            ExcelApp.Range("Q2:Q3").VerticalAlignment = xlCenter
            ExcelApp.Range("Q2:Q3").HorizontalAlignment = xlCenter


            ExcelApp.Cells(iBgnRow, 1) = "STT"
            ExcelApp.Cells(iBgnRow, 2) = "Bộ Phận"
            ExcelApp.Cells(iBgnRow, 3) = "Trợ Lý Báo Cơm"
            ExcelApp.Cells(iBgnRow, 4) = "Trưa"
            ExcelApp.Cells(iBgnRow + 1, 4) = "Cơm Mặn"
            ExcelApp.Cells(iBgnRow + 1, 5) = "Phụ Mặn"
            ExcelApp.Cells(iBgnRow + 1, 6) = "Cơm Chay"
            ExcelApp.Cells(iBgnRow + 1, 7) = "Cháo"
            ExcelApp.Cells(iBgnRow + 1, 8) = "Phụ Chay"
            ExcelApp.Cells(iBgnRow + 1, 9) = "Tổng Cộng"

            ExcelApp.Cells(iBgnRow, 10) = "Chiều"
            ExcelApp.Cells(iBgnRow + 1, 10) = "Cơm Mặn"
            ExcelApp.Cells(iBgnRow + 1, 11) = "Phụ Mặn"
            ExcelApp.Cells(iBgnRow + 1, 12) = "Phụ Chay"
            ExcelApp.Cells(iBgnRow, 13) = "Đêm"
            ExcelApp.Cells(iBgnRow + 1, 13) = "Mặn"
            ExcelApp.Cells(iBgnRow + 1, 14) = "Chay"
            ExcelApp.Cells(iBgnRow, 15) = "Cơm Khách"
            ExcelApp.Cells(iBgnRow + 1, 15) = "Dọn Bàn Riêng"
            ExcelApp.Cells(iBgnRow + 1, 16) = "Buffet"
            ExcelApp.Cells(iBgnRow, 17) = "Ghi Chú"


            ' ExcelApp.Range("A1:" + sCol + iBgnRow.ToString).VerticalAlignment = Excel.XlHAlign.xlHAlignCenter


            iBgnRow += 2
            Dim stt As Integer = 1
            For rRow As Integer = 0 To tb.Rows.Count - 1

                ExcelApp.Cells(iBgnRow, 1) = stt
                ExcelApp.Cells(iBgnRow, 2) = tb.Rows(rRow).Item("ten").ToString()
                ExcelApp.Cells(iBgnRow, 3) = tb.Rows(rRow).Item("trolydat").ToString()

                ExcelApp.Cells(iBgnRow, 4) = tb.Rows(rRow).Item("comman_trua").ToString()
                ExcelApp.Cells(iBgnRow, 5) = tb.Rows(rRow).Item("phuman_trua").ToString()
                ExcelApp.Cells(iBgnRow, 6) = tb.Rows(rRow).Item("comchay_trua").ToString()
                ExcelApp.Cells(iBgnRow, 7) = tb.Rows(rRow).Item("chao_trua").ToString()
                ExcelApp.Cells(iBgnRow, 8) = tb.Rows(rRow).Item("phuchay_trua").ToString()

                ExcelApp.Cells(iBgnRow, 9) = tb.Rows(rRow).Item("comman_trua") + tb.Rows(rRow).Item("phuman_trua") + tb.Rows(rRow).Item("comchay_trua") + tb.Rows(rRow).Item("chao_trua") + tb.Rows(rRow).Item("phuchay_trua")

                ExcelApp.Cells(iBgnRow, 10) = tb.Rows(rRow).Item("comman_chieu").ToString()
                ExcelApp.Cells(iBgnRow, 11) = tb.Rows(rRow).Item("phuman_chieu").ToString()
                ExcelApp.Cells(iBgnRow, 12) = tb.Rows(rRow).Item("phuchay_chieu").ToString()

                ExcelApp.Cells(iBgnRow, 13) = tb.Rows(rRow).Item("man_dem").ToString()
                ExcelApp.Cells(iBgnRow, 14) = tb.Rows(rRow).Item("chay_dem").ToString()

                ExcelApp.Cells(iBgnRow, 15) = tb.Rows(rRow).Item("donbanrieng").ToString()
                ExcelApp.Cells(iBgnRow, 16) = tb.Rows(rRow).Item("buffet").ToString()

                ExcelApp.Cells(iBgnRow, 17) = tb.Rows(rRow).Item("ghichu").ToString()
                stt += 1
                iBgnRow += 1
            Next
            ExcelApp.Range("I2:I" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("Q2:Q" & iBgnRow - 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown)
            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).MergeCells = True
            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Bold = True
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("D4:P" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Cells(iBgnRow, 1) = "Tổng Cột"
            ExcelApp.Cells(iBgnRow, 4) = "=SUM(D4:D" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 5) = "=SUM(E4:E" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 6) = "=SUM(F4:F" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 7) = "=SUM(G4:G" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 8) = "=SUM(H4:H" + (iBgnRow - 1).ToString + ")"


            ExcelApp.Cells(iBgnRow, 10) = "=SUM(J4:J" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 11) = "=SUM(K4:K" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 12) = "=SUM(L4:L" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 13) = "=SUM(M4:M" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 14) = "=SUM(N4:N" + (iBgnRow - 1).ToString + ")"
            iBgnRow += 1

            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).MergeCells = True
            ExcelApp.Range("A" & iBgnRow & ":C" & iBgnRow).Font.Color = Color.Red
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Bold = True
            ExcelApp.Range("A" & iBgnRow & ":P" & iBgnRow).Font.Color = Color.Red


            ExcelApp.Cells(iBgnRow, 1) = "Tổng Chung"
            ExcelApp.Range("D" & iBgnRow & ":H" & iBgnRow).MergeCells = True
            ExcelApp.Range("D" & iBgnRow & ":H" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Range("J" & iBgnRow & ":L" & iBgnRow).MergeCells = True
            ExcelApp.Range("J" & iBgnRow & ":L" & iBgnRow).HorizontalAlignment = xlCenter
            ExcelApp.Range("M" & iBgnRow & ":N" & iBgnRow).MergeCells = True
            ExcelApp.Range("M" & iBgnRow & ":N" & iBgnRow).HorizontalAlignment = xlCenter

            ExcelApp.Cells(iBgnRow, 4) = "=SUM(D" + (iBgnRow - 1).ToString + ":H" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 10) = "=SUM(J" + (iBgnRow - 1).ToString + ":L" + (iBgnRow - 1).ToString + ")"
            ExcelApp.Cells(iBgnRow, 13) = "=SUM(M" + (iBgnRow - 1).ToString + ":N" + (iBgnRow - 1).ToString + ")"
            iBgnRow += 1

            ExcelApp.Range("A2:" & sCol & (iBgnRow - 1).ToString).Borders.LineStyle = 1 'Excel.XlLineStyle.xlContinuous
            ExcelApp.Range("D5:E" + iBgnRow.ToString).WrapText = True
            ExcelApp.Range("A4:" + sCol + iBgnRow.ToString).Font.Size = 10
            ExcelApp.Range("A1:" + sCol + iBgnRow.ToString).Font.Name = "Times New Roman"

            ExcelApp.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
