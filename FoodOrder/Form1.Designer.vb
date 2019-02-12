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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Panel_title = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panelsys = New System.Windows.Forms.Panel()
        Me.Panelbody = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel_cont = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dtp_datem = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.pnl_select = New System.Windows.Forms.Panel()
        Me.rd_comkhach = New System.Windows.Forms.RadioButton()
        Me.rd_dem = New System.Windows.Forms.RadioButton()
        Me.rd_chieu = New System.Windows.Forms.RadioButton()
        Me.rd_trua = New System.Windows.Forms.RadioButton()
        Me.Panel_select = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtp_date = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txt_note = New System.Windows.Forms.TextBox()
        Me.dgv_food = New System.Windows.Forms.DataGridView()
        Me.cb_dept = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lbl_pnl_select_title = New System.Windows.Forms.Label()
        Me.btn_ex2 = New System.Windows.Forms.Button()
        Me.btn_upd = New System.Windows.Forms.Button()
        Me.btn_save = New System.Windows.Forms.Button()
        Me.btn_cancel = New System.Windows.Forms.Button()
        Me.btn_excel = New System.Windows.Forms.Button()
        Me.btn_pnl_select_ok = New System.Windows.Forms.Button()
        Me.btn_pnl_select_close = New System.Windows.Forms.Button()
        Me.btn_query = New System.Windows.Forms.Button()
        Me.btn_add = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel_title.SuspendLayout()
        Me.Panelsys.SuspendLayout()
        Me.Panelbody.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel_cont.SuspendLayout()
        Me.pnl_select.SuspendLayout()
        Me.Panel_select.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgv_food, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel_title
        '
        Me.Panel_title.Controls.Add(Me.Label1)
        Me.Panel_title.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_title.Location = New System.Drawing.Point(0, 0)
        Me.Panel_title.Name = "Panel_title"
        Me.Panel_title.Size = New System.Drawing.Size(1252, 45)
        Me.Panel_title.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(449, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(296, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "CHƯƠNG TRÌNH ĐẶT CƠM"
        '
        'Panelsys
        '
        Me.Panelsys.Controls.Add(Me.Panelbody)
        Me.Panelsys.Controls.Add(Me.Panel_title)
        Me.Panelsys.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panelsys.Location = New System.Drawing.Point(0, 0)
        Me.Panelsys.Name = "Panelsys"
        Me.Panelsys.Size = New System.Drawing.Size(1252, 714)
        Me.Panelsys.TabIndex = 32
        '
        'Panelbody
        '
        Me.Panelbody.Controls.Add(Me.SplitContainer1)
        Me.Panelbody.Location = New System.Drawing.Point(34, 83)
        Me.Panelbody.Name = "Panelbody"
        Me.Panelbody.Size = New System.Drawing.Size(1210, 409)
        Me.Panelbody.TabIndex = 3
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel_cont)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.PictureBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1210, 409)
        Me.SplitContainer1.SplitterDistance = 729
        Me.SplitContainer1.TabIndex = 2
        '
        'Panel_cont
        '
        Me.Panel_cont.Controls.Add(Me.Label6)
        Me.Panel_cont.Controls.Add(Me.dtp_datem)
        Me.Panel_cont.Controls.Add(Me.btn_ex2)
        Me.Panel_cont.Controls.Add(Me.btn_upd)
        Me.Panel_cont.Controls.Add(Me.Label4)
        Me.Panel_cont.Controls.Add(Me.btn_save)
        Me.Panel_cont.Controls.Add(Me.Label2)
        Me.Panel_cont.Controls.Add(Me.btn_cancel)
        Me.Panel_cont.Controls.Add(Me.btn_excel)
        Me.Panel_cont.Controls.Add(Me.pnl_select)
        Me.Panel_cont.Controls.Add(Me.dtp_date)
        Me.Panel_cont.Controls.Add(Me.GroupBox1)
        Me.Panel_cont.Controls.Add(Me.btn_query)
        Me.Panel_cont.Controls.Add(Me.dgv_food)
        Me.Panel_cont.Controls.Add(Me.btn_add)
        Me.Panel_cont.Controls.Add(Me.cb_dept)
        Me.Panel_cont.Controls.Add(Me.GroupBox2)
        Me.Panel_cont.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_cont.Location = New System.Drawing.Point(0, 0)
        Me.Panel_cont.Name = "Panel_cont"
        Me.Panel_cont.Size = New System.Drawing.Size(729, 409)
        Me.Panel_cont.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(374, 440)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Tháng: "
        Me.Label6.Visible = False
        '
        'dtp_datem
        '
        Me.dtp_datem.CustomFormat = "MM"
        Me.dtp_datem.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_datem.Location = New System.Drawing.Point(424, 434)
        Me.dtp_datem.Name = "dtp_datem"
        Me.dtp_datem.Size = New System.Drawing.Size(62, 20)
        Me.dtp_datem.TabIndex = 30
        Me.dtp_datem.Value = New Date(2018, 3, 7, 0, 0, 0, 0)
        Me.dtp_datem.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(192, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Ngày Đặt:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Bộ Phận:"
        '
        'pnl_select
        '
        Me.pnl_select.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnl_select.Controls.Add(Me.btn_pnl_select_ok)
        Me.pnl_select.Controls.Add(Me.rd_comkhach)
        Me.pnl_select.Controls.Add(Me.rd_dem)
        Me.pnl_select.Controls.Add(Me.rd_chieu)
        Me.pnl_select.Controls.Add(Me.rd_trua)
        Me.pnl_select.Controls.Add(Me.Panel_select)
        Me.pnl_select.Location = New System.Drawing.Point(261, 88)
        Me.pnl_select.Name = "pnl_select"
        Me.pnl_select.Size = New System.Drawing.Size(244, 167)
        Me.pnl_select.TabIndex = 24
        Me.pnl_select.Visible = False
        '
        'rd_comkhach
        '
        Me.rd_comkhach.AutoSize = True
        Me.rd_comkhach.ForeColor = System.Drawing.Color.Blue
        Me.rd_comkhach.Location = New System.Drawing.Point(29, 104)
        Me.rd_comkhach.Name = "rd_comkhach"
        Me.rd_comkhach.Size = New System.Drawing.Size(80, 17)
        Me.rd_comkhach.TabIndex = 16
        Me.rd_comkhach.TabStop = True
        Me.rd_comkhach.Text = "Cơm Khách"
        Me.rd_comkhach.UseVisualStyleBackColor = True
        '
        'rd_dem
        '
        Me.rd_dem.AutoSize = True
        Me.rd_dem.ForeColor = System.Drawing.Color.Blue
        Me.rd_dem.Location = New System.Drawing.Point(29, 81)
        Me.rd_dem.Name = "rd_dem"
        Me.rd_dem.Size = New System.Drawing.Size(47, 17)
        Me.rd_dem.TabIndex = 15
        Me.rd_dem.TabStop = True
        Me.rd_dem.Text = "Đêm"
        Me.rd_dem.UseVisualStyleBackColor = True
        '
        'rd_chieu
        '
        Me.rd_chieu.AutoSize = True
        Me.rd_chieu.ForeColor = System.Drawing.Color.Blue
        Me.rd_chieu.Location = New System.Drawing.Point(29, 58)
        Me.rd_chieu.Name = "rd_chieu"
        Me.rd_chieu.Size = New System.Drawing.Size(52, 17)
        Me.rd_chieu.TabIndex = 14
        Me.rd_chieu.TabStop = True
        Me.rd_chieu.Text = "Chiều"
        Me.rd_chieu.UseVisualStyleBackColor = True
        '
        'rd_trua
        '
        Me.rd_trua.AutoSize = True
        Me.rd_trua.Checked = True
        Me.rd_trua.ForeColor = System.Drawing.Color.Blue
        Me.rd_trua.Location = New System.Drawing.Point(29, 35)
        Me.rd_trua.Name = "rd_trua"
        Me.rd_trua.Size = New System.Drawing.Size(47, 17)
        Me.rd_trua.TabIndex = 13
        Me.rd_trua.TabStop = True
        Me.rd_trua.Text = "Trưa"
        Me.rd_trua.UseVisualStyleBackColor = True
        '
        'Panel_select
        '
        Me.Panel_select.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(84, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.Panel_select.Controls.Add(Me.Label3)
        Me.Panel_select.Controls.Add(Me.btn_pnl_select_close)
        Me.Panel_select.Location = New System.Drawing.Point(0, 0)
        Me.Panel_select.Name = "Panel_select"
        Me.Panel_select.Size = New System.Drawing.Size(242, 24)
        Me.Panel_select.TabIndex = 12
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(75, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 19)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Suất Ăn"
        '
        'dtp_date
        '
        Me.dtp_date.CustomFormat = "yyyy/MM/dd"
        Me.dtp_date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtp_date.Location = New System.Drawing.Point(253, 31)
        Me.dtp_date.Name = "dtp_date"
        Me.dtp_date.Size = New System.Drawing.Size(120, 20)
        Me.dtp_date.TabIndex = 7
        Me.dtp_date.Value = New Date(2018, 3, 7, 0, 0, 0, 0)
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txt_note)
        Me.GroupBox1.ForeColor = System.Drawing.Color.Red
        Me.GroupBox1.Location = New System.Drawing.Point(3, 340)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(715, 84)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Ghi chú"
        '
        'txt_note
        '
        Me.txt_note.Enabled = False
        Me.txt_note.Location = New System.Drawing.Point(0, 16)
        Me.txt_note.Multiline = True
        Me.txt_note.Name = "txt_note"
        Me.txt_note.Size = New System.Drawing.Size(715, 47)
        Me.txt_note.TabIndex = 0
        Me.txt_note.Text = resources.GetString("txt_note.Text")
        '
        'dgv_food
        '
        Me.dgv_food.AllowUserToAddRows = False
        Me.dgv_food.AllowUserToDeleteRows = False
        Me.dgv_food.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_food.Location = New System.Drawing.Point(0, 68)
        Me.dgv_food.Name = "dgv_food"
        Me.dgv_food.Size = New System.Drawing.Size(721, 266)
        Me.dgv_food.TabIndex = 25
        '
        'cb_dept
        '
        Me.cb_dept.FormattingEnabled = True
        Me.cb_dept.Location = New System.Drawing.Point(60, 31)
        Me.cb_dept.Name = "cb_dept"
        Me.cb_dept.Size = New System.Drawing.Size(126, 21)
        Me.cb_dept.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.ForeColor = System.Drawing.Color.Red
        Me.GroupBox2.Location = New System.Drawing.Point(2, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(719, 56)
        Me.GroupBox2.TabIndex = 28
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Điều kiện"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Fuchsia
        Me.Label5.Location = New System.Drawing.Point(199, 3)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 20)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Thực Đơn"
        '
        'lbl_pnl_select_title
        '
        Me.lbl_pnl_select_title.AutoSize = True
        Me.lbl_pnl_select_title.ForeColor = System.Drawing.Color.Red
        Me.lbl_pnl_select_title.Location = New System.Drawing.Point(7, 89)
        Me.lbl_pnl_select_title.Name = "lbl_pnl_select_title"
        Me.lbl_pnl_select_title.Size = New System.Drawing.Size(0, 13)
        Me.lbl_pnl_select_title.TabIndex = 31
        '
        'btn_ex2
        '
        Me.btn_ex2.Image = Global.FoodOrder.My.Resources.Resources.print
        Me.btn_ex2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_ex2.Location = New System.Drawing.Point(564, 426)
        Me.btn_ex2.Name = "btn_ex2"
        Me.btn_ex2.Size = New System.Drawing.Size(140, 40)
        Me.btn_ex2.TabIndex = 29
        Me.btn_ex2.Text = "  In Excel Theo Tháng"
        Me.btn_ex2.UseVisualStyleBackColor = True
        Me.btn_ex2.Visible = False
        '
        'btn_upd
        '
        Me.btn_upd.Image = Global.FoodOrder.My.Resources.Resources.edit
        Me.btn_upd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_upd.Location = New System.Drawing.Point(481, 31)
        Me.btn_upd.Name = "btn_upd"
        Me.btn_upd.Size = New System.Drawing.Size(77, 33)
        Me.btn_upd.TabIndex = 10
        Me.btn_upd.Text = "Sửa"
        Me.btn_upd.UseVisualStyleBackColor = True
        '
        'btn_save
        '
        Me.btn_save.Image = Global.FoodOrder.My.Resources.Resources.save
        Me.btn_save.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_save.Location = New System.Drawing.Point(138, 430)
        Me.btn_save.Name = "btn_save"
        Me.btn_save.Size = New System.Drawing.Size(77, 33)
        Me.btn_save.TabIndex = 26
        Me.btn_save.Text = "Lưu"
        Me.btn_save.UseVisualStyleBackColor = True
        '
        'btn_cancel
        '
        Me.btn_cancel.Image = Global.FoodOrder.My.Resources.Resources.close
        Me.btn_cancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_cancel.Location = New System.Drawing.Point(235, 430)
        Me.btn_cancel.Name = "btn_cancel"
        Me.btn_cancel.Size = New System.Drawing.Size(77, 33)
        Me.btn_cancel.TabIndex = 27
        Me.btn_cancel.Text = "Hủy"
        Me.btn_cancel.UseVisualStyleBackColor = True
        '
        'btn_excel
        '
        Me.btn_excel.Image = Global.FoodOrder.My.Resources.Resources.print
        Me.btn_excel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_excel.Location = New System.Drawing.Point(644, 31)
        Me.btn_excel.Name = "btn_excel"
        Me.btn_excel.Size = New System.Drawing.Size(77, 33)
        Me.btn_excel.TabIndex = 11
        Me.btn_excel.Text = "  In Excel"
        Me.btn_excel.UseVisualStyleBackColor = True
        '
        'btn_pnl_select_ok
        '
        Me.btn_pnl_select_ok.Image = Global.FoodOrder.My.Resources.Resources.OK
        Me.btn_pnl_select_ok.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_pnl_select_ok.Location = New System.Drawing.Point(79, 127)
        Me.btn_pnl_select_ok.Name = "btn_pnl_select_ok"
        Me.btn_pnl_select_ok.Size = New System.Drawing.Size(77, 33)
        Me.btn_pnl_select_ok.TabIndex = 17
        Me.btn_pnl_select_ok.Text = "Chọn"
        Me.btn_pnl_select_ok.UseVisualStyleBackColor = True
        '
        'btn_pnl_select_close
        '
        Me.btn_pnl_select_close.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_pnl_select_close.Image = CType(resources.GetObject("btn_pnl_select_close.Image"), System.Drawing.Image)
        Me.btn_pnl_select_close.Location = New System.Drawing.Point(219, 0)
        Me.btn_pnl_select_close.Name = "btn_pnl_select_close"
        Me.btn_pnl_select_close.Size = New System.Drawing.Size(23, 24)
        Me.btn_pnl_select_close.TabIndex = 0
        Me.btn_pnl_select_close.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_pnl_select_close.UseVisualStyleBackColor = True
        '
        'btn_query
        '
        Me.btn_query.Image = Global.FoodOrder.My.Resources.Resources.Find
        Me.btn_query.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_query.Location = New System.Drawing.Point(564, 31)
        Me.btn_query.Name = "btn_query"
        Me.btn_query.Size = New System.Drawing.Size(77, 33)
        Me.btn_query.TabIndex = 8
        Me.btn_query.Text = "Tra Cứu"
        Me.btn_query.UseVisualStyleBackColor = True
        '
        'btn_add
        '
        Me.btn_add.Image = Global.FoodOrder.My.Resources.Resources.Add
        Me.btn_add.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_add.Location = New System.Drawing.Point(398, 31)
        Me.btn_add.Name = "btn_add"
        Me.btn_add.Size = New System.Drawing.Size(77, 33)
        Me.btn_add.TabIndex = 9
        Me.btn_add.Text = "Thêm"
        Me.btn_add.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(477, 409)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1252, 714)
        Me.Controls.Add(Me.Panelsys)
        Me.Controls.Add(Me.lbl_pnl_select_title)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.Panel_title.ResumeLayout(False)
        Me.Panel_title.PerformLayout()
        Me.Panelsys.ResumeLayout(False)
        Me.Panelbody.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel_cont.ResumeLayout(False)
        Me.Panel_cont.PerformLayout()
        Me.pnl_select.ResumeLayout(False)
        Me.pnl_select.PerformLayout()
        Me.Panel_select.ResumeLayout(False)
        Me.Panel_select.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgv_food, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel_title As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panelsys As System.Windows.Forms.Panel
    Friend WithEvents Panelbody As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel_cont As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtp_datem As System.Windows.Forms.DateTimePicker
    Friend WithEvents btn_ex2 As System.Windows.Forms.Button
    Friend WithEvents btn_upd As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btn_save As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btn_cancel As System.Windows.Forms.Button
    Friend WithEvents btn_excel As System.Windows.Forms.Button
    Friend WithEvents pnl_select As System.Windows.Forms.Panel
    Friend WithEvents btn_pnl_select_ok As System.Windows.Forms.Button
    Friend WithEvents rd_comkhach As System.Windows.Forms.RadioButton
    Friend WithEvents rd_dem As System.Windows.Forms.RadioButton
    Friend WithEvents rd_chieu As System.Windows.Forms.RadioButton
    Friend WithEvents rd_trua As System.Windows.Forms.RadioButton
    Friend WithEvents Panel_select As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btn_pnl_select_close As System.Windows.Forms.Button
    Friend WithEvents dtp_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txt_note As System.Windows.Forms.TextBox
    Friend WithEvents btn_query As System.Windows.Forms.Button
    Friend WithEvents dgv_food As System.Windows.Forms.DataGridView
    Friend WithEvents btn_add As System.Windows.Forms.Button
    Friend WithEvents cb_dept As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lbl_pnl_select_title As System.Windows.Forms.Label
End Class
