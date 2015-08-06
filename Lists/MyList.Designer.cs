namespace Priem
{
    partial class MyList
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.cbStudyBasis = new System.Windows.Forms.ComboBox();
            this.cbStudyForm = new System.Windows.Forms.ComboBox();
            this.cbLicenseProgram = new System.Windows.Forms.ComboBox();
            this.btnFillGrid = new System.Windows.Forms.Button();
            this.dgvAbitList = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbZeroWave = new System.Windows.Forms.CheckBox();
            this.tbAbitsTop = new System.Windows.Forms.TextBox();
            this.rbAbitsTop = new System.Windows.Forms.RadioButton();
            this.rbAbitsAll = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label8 = new System.Windows.Forms.Label();
            this.pictureBoxWhite = new System.Windows.Forms.PictureBox();
            this.labelWhite = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.btn_GreenList = new System.Windows.Forms.Button();
            this.pictureBoxThistle = new System.Windows.Forms.PictureBox();
            this.labelThistle = new System.Windows.Forms.Label();
            this.ttextEntryView = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.rbCrimea = new System.Windows.Forms.RadioButton();
            this.rbForeign = new System.Windows.Forms.RadioButton();
            this.rbCommon = new System.Windows.Forms.RadioButton();
            this.btnRePaint = new System.Windows.Forms.Button();
            this.pictureBoxBeige = new System.Windows.Forms.PictureBox();
            this.labelBeige = new System.Windows.Forms.Label();
            this.btnRestoreOriginals = new System.Windows.Forms.Button();
            this.gb1kurs = new System.Windows.Forms.GroupBox();
            this.btnSetAllOrigins = new System.Windows.Forms.Button();
            this.chbOnlyWithOrigins = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.tbDinamicWave = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.gbMag = new System.Windows.Forms.GroupBox();
            this.chbOnlyWithOriginsMag = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.tbDinamicWaveMag = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAbitList)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxWhite)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxThistle)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxBeige)).BeginInit();
            this.gb1kurs.SuspendLayout();
            this.gbMag.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblCount
            // 
            this.lblCount.Location = new System.Drawing.Point(42, 523);
            this.lblCount.Visible = false;
            // 
            // btnCard
            // 
            this.btnCard.Location = new System.Drawing.Point(245, 515);
            this.btnCard.Visible = false;
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(13, 518);
            this.btnRemove.Visible = false;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(13, 521);
            this.btnAdd.Visible = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(231, 515);
            this.btnClose.Visible = false;
            // 
            // cbFaculty
            // 
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(102, 18);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(194, 21);
            this.cbFaculty.TabIndex = 0;
            this.cbFaculty.SelectedIndexChanged += new System.EventHandler(this.cbFaculty_SelectedIndexChanged);
            // 
            // cbStudyBasis
            // 
            this.cbStudyBasis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyBasis.FormattingEnabled = true;
            this.cbStudyBasis.Location = new System.Drawing.Point(102, 17);
            this.cbStudyBasis.Name = "cbStudyBasis";
            this.cbStudyBasis.Size = new System.Drawing.Size(194, 21);
            this.cbStudyBasis.TabIndex = 1;
            this.cbStudyBasis.SelectedIndexChanged += new System.EventHandler(this.cbStudyBasis_SelectedIndexChanged);
            // 
            // cbStudyForm
            // 
            this.cbStudyForm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyForm.FormattingEnabled = true;
            this.cbStudyForm.Location = new System.Drawing.Point(102, 44);
            this.cbStudyForm.Name = "cbStudyForm";
            this.cbStudyForm.Size = new System.Drawing.Size(194, 21);
            this.cbStudyForm.TabIndex = 2;
            // 
            // cbLicenseProgram
            // 
            this.cbLicenseProgram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLicenseProgram.FormattingEnabled = true;
            this.cbLicenseProgram.Location = new System.Drawing.Point(102, 45);
            this.cbLicenseProgram.Name = "cbLicenseProgram";
            this.cbLicenseProgram.Size = new System.Drawing.Size(194, 21);
            this.cbLicenseProgram.TabIndex = 3;
            this.cbLicenseProgram.SelectedIndexChanged += new System.EventHandler(this.cbLicenseProgram_SelectedIndexChanged);
            // 
            // btnFillGrid
            // 
            this.btnFillGrid.Location = new System.Drawing.Point(342, 116);
            this.btnFillGrid.Name = "btnFillGrid";
            this.btnFillGrid.Size = new System.Drawing.Size(189, 23);
            this.btnFillGrid.TabIndex = 4;
            this.btnFillGrid.Text = "Запустить подсчёт";
            this.btnFillGrid.UseVisualStyleBackColor = true;
            this.btnFillGrid.Click += new System.EventHandler(this.btnFillGrid_Click);
            // 
            // dgvAbitList
            // 
            this.dgvAbitList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvAbitList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAbitList.Location = new System.Drawing.Point(13, 174);
            this.dgvAbitList.Name = "dgvAbitList";
            this.dgvAbitList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvAbitList.Size = new System.Drawing.Size(1293, 373);
            this.dgvAbitList.TabIndex = 5;
            this.dgvAbitList.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvAbitList_CellMouseClick);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbZeroWave);
            this.groupBox2.Controls.Add(this.tbAbitsTop);
            this.groupBox2.Controls.Add(this.rbAbitsTop);
            this.groupBox2.Controls.Add(this.rbAbitsAll);
            this.groupBox2.Location = new System.Drawing.Point(331, 9);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(320, 61);
            this.groupBox2.TabIndex = 67;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Количество абитуриентов для отображения";
            // 
            // cbZeroWave
            // 
            this.cbZeroWave.AutoSize = true;
            this.cbZeroWave.Location = new System.Drawing.Point(11, 39);
            this.cbZeroWave.Name = "cbZeroWave";
            this.cbZeroWave.Size = new System.Drawing.Size(183, 17);
            this.cbZeroWave.TabIndex = 86;
            this.cbZeroWave.Text = "Выводить только зачисленных";
            this.cbZeroWave.UseVisualStyleBackColor = true;
            // 
            // tbAbitsTop
            // 
            this.tbAbitsTop.Location = new System.Drawing.Point(261, 16);
            this.tbAbitsTop.Name = "tbAbitsTop";
            this.tbAbitsTop.Size = new System.Drawing.Size(53, 20);
            this.tbAbitsTop.TabIndex = 2;
            this.tbAbitsTop.MouseClick += new System.Windows.Forms.MouseEventHandler(this.tbAbitsTop_MouseClick);
            // 
            // rbAbitsTop
            // 
            this.rbAbitsTop.AutoSize = true;
            this.rbAbitsTop.Location = new System.Drawing.Point(130, 17);
            this.rbAbitsTop.Name = "rbAbitsTop";
            this.rbAbitsTop.Size = new System.Drawing.Size(128, 17);
            this.rbAbitsTop.TabIndex = 1;
            this.rbAbitsTop.TabStop = true;
            this.rbAbitsTop.Text = "Отображать первые";
            this.rbAbitsTop.UseVisualStyleBackColor = true;
            // 
            // rbAbitsAll
            // 
            this.rbAbitsAll.AutoSize = true;
            this.rbAbitsAll.Location = new System.Drawing.Point(11, 17);
            this.rbAbitsAll.Name = "rbAbitsAll";
            this.rbAbitsAll.Size = new System.Drawing.Size(113, 17);
            this.rbAbitsAll.TabIndex = 0;
            this.rbAbitsAll.TabStop = true;
            this.rbAbitsAll.Text = "Отображать всех";
            this.rbAbitsAll.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 68;
            this.label1.Text = "Факультет";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(94, 13);
            this.label2.TabIndex = 69;
            this.label2.Text = "Основа обучения";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 70;
            this.label3.Text = "Форма обучения";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 71;
            this.label4.Text = "Направление";
            // 
            // btnToExcel
            // 
            this.btnToExcel.Location = new System.Drawing.Point(342, 145);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(189, 21);
            this.btnToExcel.TabIndex = 72;
            this.btnToExcel.Text = "Распечатать в Excel";
            this.btnToExcel.UseVisualStyleBackColor = true;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(690, 6);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(396, 13);
            this.label5.TabIndex = 73;
            this.label5.Text = "Абитуриент рекомендован к зачислению на более приоритетную программу";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(690, 23);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(329, 13);
            this.label6.TabIndex = 74;
            this.label6.Text = "Абитуриент рекомендован к зачислению на данную программу\r\n";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(690, 40);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(261, 13);
            this.label7.TabIndex = 75;
            this.label7.Text = "Отсутствует приоритет или внутренний приоритет";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Yellow;
            this.pictureBox1.Location = new System.Drawing.Point(664, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(25, 15);
            this.pictureBox1.TabIndex = 76;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.LightGreen;
            this.pictureBox2.Location = new System.Drawing.Point(664, 22);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(25, 15);
            this.pictureBox2.TabIndex = 77;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackColor = System.Drawing.Color.LightBlue;
            this.pictureBox3.Location = new System.Drawing.Point(664, 39);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(25, 15);
            this.pictureBox3.TabIndex = 78;
            this.pictureBox3.TabStop = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(664, 111);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(422, 13);
            this.label8.TabIndex = 79;
            this.label8.Text = "Ид.Номер_ФИО (Приоритет_программы, внутренний_приоритет, сумма_баллов)\r\n";
            // 
            // pictureBoxWhite
            // 
            this.pictureBoxWhite.BackColor = System.Drawing.Color.White;
            this.pictureBoxWhite.Location = new System.Drawing.Point(664, 73);
            this.pictureBoxWhite.Name = "pictureBoxWhite";
            this.pictureBoxWhite.Size = new System.Drawing.Size(25, 15);
            this.pictureBoxWhite.TabIndex = 81;
            this.pictureBoxWhite.TabStop = false;
            // 
            // labelWhite
            // 
            this.labelWhite.AutoSize = true;
            this.labelWhite.Location = new System.Drawing.Point(690, 74);
            this.labelWhite.Name = "labelWhite";
            this.labelWhite.Size = new System.Drawing.Size(367, 13);
            this.labelWhite.TabIndex = 80;
            this.labelWhite.Text = "Не рекомендован к зачислению на эту программу (или нет оригинала)";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label10.Location = new System.Drawing.Point(661, 124);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(413, 15);
            this.label10.TabIndex = 82;
            this.label10.Text = "Щелкните правой мышкой по направлению или абитуриенту";
            // 
            // btn_GreenList
            // 
            this.btn_GreenList.Location = new System.Drawing.Point(537, 116);
            this.btn_GreenList.Name = "btn_GreenList";
            this.btn_GreenList.Size = new System.Drawing.Size(108, 49);
            this.btn_GreenList.TabIndex = 83;
            this.btn_GreenList.Text = "GreenList";
            this.btn_GreenList.UseVisualStyleBackColor = true;
            // 
            // pictureBoxThistle
            // 
            this.pictureBoxThistle.BackColor = System.Drawing.Color.Thistle;
            this.pictureBoxThistle.Location = new System.Drawing.Point(664, 56);
            this.pictureBoxThistle.Name = "pictureBoxThistle";
            this.pictureBoxThistle.Size = new System.Drawing.Size(25, 15);
            this.pictureBoxThistle.TabIndex = 85;
            this.pictureBoxThistle.TabStop = false;
            // 
            // labelThistle
            // 
            this.labelThistle.AutoSize = true;
            this.labelThistle.Location = new System.Drawing.Point(690, 58);
            this.labelThistle.Name = "labelThistle";
            this.labelThistle.Size = new System.Drawing.Size(374, 13);
            this.labelThistle.TabIndex = 84;
            this.labelThistle.Text = "У абитуриента нет необходимого ЕГЭ для образовательной программы";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.cbStudyForm);
            this.groupBox1.Controls.Add(this.cbStudyBasis);
            this.groupBox1.Location = new System.Drawing.Point(13, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(310, 79);
            this.groupBox1.TabIndex = 86;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Учитывается при анализе:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.cbLicenseProgram);
            this.groupBox3.Controls.Add(this.cbFaculty);
            this.groupBox3.Location = new System.Drawing.Point(13, 90);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(310, 78);
            this.groupBox3.TabIndex = 87;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Учитывается при отображении:";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.rbCrimea);
            this.groupBox4.Controls.Add(this.rbForeign);
            this.groupBox4.Controls.Add(this.rbCommon);
            this.groupBox4.Location = new System.Drawing.Point(331, 74);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(320, 37);
            this.groupBox4.TabIndex = 88;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Конкурс";
            // 
            // rbCrimea
            // 
            this.rbCrimea.AutoSize = true;
            this.rbCrimea.Location = new System.Drawing.Point(243, 14);
            this.rbCrimea.Name = "rbCrimea";
            this.rbCrimea.Size = new System.Drawing.Size(53, 17);
            this.rbCrimea.TabIndex = 2;
            this.rbCrimea.Text = "крым";
            this.rbCrimea.UseVisualStyleBackColor = true;
            // 
            // rbForeign
            // 
            this.rbForeign.AutoSize = true;
            this.rbForeign.Location = new System.Drawing.Point(130, 14);
            this.rbForeign.Name = "rbForeign";
            this.rbForeign.Size = new System.Drawing.Size(63, 17);
            this.rbForeign.TabIndex = 1;
            this.rbForeign.Text = "иностр.";
            this.rbForeign.UseVisualStyleBackColor = true;
            // 
            // rbCommon
            // 
            this.rbCommon.AutoSize = true;
            this.rbCommon.Checked = true;
            this.rbCommon.Location = new System.Drawing.Point(11, 14);
            this.rbCommon.Name = "rbCommon";
            this.rbCommon.Size = new System.Drawing.Size(58, 17);
            this.rbCommon.TabIndex = 0;
            this.rbCommon.TabStop = true;
            this.rbCommon.Text = "общий";
            this.rbCommon.UseVisualStyleBackColor = true;
            // 
            // btnRePaint
            // 
            this.btnRePaint.Location = new System.Drawing.Point(6, 18);
            this.btnRePaint.Name = "btnRePaint";
            this.btnRePaint.Size = new System.Drawing.Size(199, 23);
            this.btnRePaint.TabIndex = 89;
            this.btnRePaint.Text = "Пересчитать";
            this.btnRePaint.UseVisualStyleBackColor = true;
            this.btnRePaint.Click += new System.EventHandler(this.btnRePaint_Click);
            // 
            // pictureBoxBeige
            // 
            this.pictureBoxBeige.BackColor = System.Drawing.Color.Beige;
            this.pictureBoxBeige.Location = new System.Drawing.Point(664, 90);
            this.pictureBoxBeige.Name = "pictureBoxBeige";
            this.pictureBoxBeige.Size = new System.Drawing.Size(25, 15);
            this.pictureBoxBeige.TabIndex = 91;
            this.pictureBoxBeige.TabStop = false;
            // 
            // labelBeige
            // 
            this.labelBeige.AutoSize = true;
            this.labelBeige.Location = new System.Drawing.Point(690, 91);
            this.labelBeige.Name = "labelBeige";
            this.labelBeige.Size = new System.Drawing.Size(214, 13);
            this.labelBeige.TabIndex = 90;
            this.labelBeige.Text = "Находится в 80%-зоне, но нет оригинала";
            // 
            // btnRestoreOriginals
            // 
            this.btnRestoreOriginals.Location = new System.Drawing.Point(6, 42);
            this.btnRestoreOriginals.Name = "btnRestoreOriginals";
            this.btnRestoreOriginals.Size = new System.Drawing.Size(199, 24);
            this.btnRestoreOriginals.TabIndex = 92;
            this.btnRestoreOriginals.Text = "Восстановить значения оригиналов";
            this.btnRestoreOriginals.UseVisualStyleBackColor = true;
            this.btnRestoreOriginals.Click += new System.EventHandler(this.btnRestoreOriginals_Click);
            // 
            // gb1kurs
            // 
            this.gb1kurs.Controls.Add(this.btnSetAllOrigins);
            this.gb1kurs.Controls.Add(this.chbOnlyWithOrigins);
            this.gb1kurs.Controls.Add(this.label11);
            this.gb1kurs.Controls.Add(this.tbDinamicWave);
            this.gb1kurs.Controls.Add(this.label9);
            this.gb1kurs.Controls.Add(this.btnRePaint);
            this.gb1kurs.Controls.Add(this.btnRestoreOriginals);
            this.gb1kurs.Location = new System.Drawing.Point(1095, 5);
            this.gb1kurs.Name = "gb1kurs";
            this.gb1kurs.Size = new System.Drawing.Size(211, 162);
            this.gb1kurs.TabIndex = 93;
            this.gb1kurs.TabStop = false;
            this.gb1kurs.Text = "1 курс";
            // 
            // btnSetAllOrigins
            // 
            this.btnSetAllOrigins.Location = new System.Drawing.Point(6, 69);
            this.btnSetAllOrigins.Name = "btnSetAllOrigins";
            this.btnSetAllOrigins.Size = new System.Drawing.Size(199, 24);
            this.btnSetAllOrigins.TabIndex = 101;
            this.btnSetAllOrigins.Text = "Проставить оригиналы всем";
            this.btnSetAllOrigins.UseVisualStyleBackColor = true;
            this.btnSetAllOrigins.Click += new System.EventHandler(this.btnSetAllOrigins_Click);
            // 
            // chbOnlyWithOrigins
            // 
            this.chbOnlyWithOrigins.Checked = true;
            this.chbOnlyWithOrigins.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chbOnlyWithOrigins.Location = new System.Drawing.Point(6, 134);
            this.chbOnlyWithOrigins.Name = "chbOnlyWithOrigins";
            this.chbOnlyWithOrigins.Size = new System.Drawing.Size(199, 22);
            this.chbOnlyWithOrigins.TabIndex = 100;
            this.chbOnlyWithOrigins.Text = "Только с оригиналами";
            this.chbOnlyWithOrigins.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(106, 114);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(15, 13);
            this.label11.TabIndex = 95;
            this.label11.Text = "%";
            // 
            // tbDinamicWave
            // 
            this.tbDinamicWave.Location = new System.Drawing.Point(6, 111);
            this.tbDinamicWave.Name = "tbDinamicWave";
            this.tbDinamicWave.Size = new System.Drawing.Size(94, 20);
            this.tbDinamicWave.TabIndex = 94;
            this.tbDinamicWave.Text = "80";
            this.tbDinamicWave.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.tbDinamicWave.TextChanged += new System.EventHandler(this.tbDinamicWave_TextChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 95);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(193, 13);
            this.label9.TabIndex = 93;
            this.label9.Text = "Кол-во зачисляемых в первую волну";
            // 
            // gbMag
            // 
            this.gbMag.Controls.Add(this.chbOnlyWithOriginsMag);
            this.gbMag.Controls.Add(this.label12);
            this.gbMag.Controls.Add(this.tbDinamicWaveMag);
            this.gbMag.Controls.Add(this.label13);
            this.gbMag.Location = new System.Drawing.Point(1095, 5);
            this.gbMag.Name = "gbMag";
            this.gbMag.Size = new System.Drawing.Size(211, 162);
            this.gbMag.TabIndex = 94;
            this.gbMag.TabStop = false;
            this.gbMag.Text = "Магистратура";
            // 
            // chbOnlyWithOriginsMag
            // 
            this.chbOnlyWithOriginsMag.Enabled = false;
            this.chbOnlyWithOriginsMag.Location = new System.Drawing.Point(6, 54);
            this.chbOnlyWithOriginsMag.Name = "chbOnlyWithOriginsMag";
            this.chbOnlyWithOriginsMag.Size = new System.Drawing.Size(199, 25);
            this.chbOnlyWithOriginsMag.TabIndex = 99;
            this.chbOnlyWithOriginsMag.Text = "Только с оригиналами";
            this.chbOnlyWithOriginsMag.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(106, 35);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(15, 13);
            this.label12.TabIndex = 98;
            this.label12.Text = "%";
            // 
            // tbDinamicWaveMag
            // 
            this.tbDinamicWaveMag.Location = new System.Drawing.Point(6, 32);
            this.tbDinamicWaveMag.Name = "tbDinamicWaveMag";
            this.tbDinamicWaveMag.ReadOnly = true;
            this.tbDinamicWaveMag.Size = new System.Drawing.Size(94, 20);
            this.tbDinamicWaveMag.TabIndex = 97;
            this.tbDinamicWaveMag.Text = "100";
            this.tbDinamicWaveMag.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(3, 17);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(193, 13);
            this.label13.TabIndex = 96;
            this.label13.Text = "Кол-во зачисляемых в первую волну";
            // 
            // NewMyList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1318, 553);
            this.Controls.Add(this.gbMag);
            this.Controls.Add(this.gb1kurs);
            this.Controls.Add(this.pictureBoxBeige);
            this.Controls.Add(this.labelBeige);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.pictureBoxThistle);
            this.Controls.Add(this.labelThistle);
            this.Controls.Add(this.btn_GreenList);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.pictureBoxWhite);
            this.Controls.Add(this.labelWhite);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnToExcel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.dgvAbitList);
            this.Controls.Add(this.btnFillGrid);
            this.Name = "NewMyList";
            this.Text = "MyList";
            this.Controls.SetChildIndex(this.btnFillGrid, 0);
            this.Controls.SetChildIndex(this.dgvAbitList, 0);
            this.Controls.SetChildIndex(this.groupBox2, 0);
            this.Controls.SetChildIndex(this.lblCount, 0);
            this.Controls.SetChildIndex(this.btnCard, 0);
            this.Controls.SetChildIndex(this.btnClose, 0);
            this.Controls.SetChildIndex(this.btnAdd, 0);
            this.Controls.SetChildIndex(this.btnRemove, 0);
            this.Controls.SetChildIndex(this.btnToExcel, 0);
            this.Controls.SetChildIndex(this.label5, 0);
            this.Controls.SetChildIndex(this.label6, 0);
            this.Controls.SetChildIndex(this.label7, 0);
            this.Controls.SetChildIndex(this.pictureBox1, 0);
            this.Controls.SetChildIndex(this.pictureBox2, 0);
            this.Controls.SetChildIndex(this.pictureBox3, 0);
            this.Controls.SetChildIndex(this.label8, 0);
            this.Controls.SetChildIndex(this.labelWhite, 0);
            this.Controls.SetChildIndex(this.pictureBoxWhite, 0);
            this.Controls.SetChildIndex(this.label10, 0);
            this.Controls.SetChildIndex(this.btn_GreenList, 0);
            this.Controls.SetChildIndex(this.labelThistle, 0);
            this.Controls.SetChildIndex(this.pictureBoxThistle, 0);
            this.Controls.SetChildIndex(this.groupBox1, 0);
            this.Controls.SetChildIndex(this.groupBox3, 0);
            this.Controls.SetChildIndex(this.groupBox4, 0);
            this.Controls.SetChildIndex(this.labelBeige, 0);
            this.Controls.SetChildIndex(this.pictureBoxBeige, 0);
            this.Controls.SetChildIndex(this.gb1kurs, 0);
            this.Controls.SetChildIndex(this.gbMag, 0);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAbitList)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxWhite)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxThistle)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxBeige)).EndInit();
            this.gb1kurs.ResumeLayout(false);
            this.gb1kurs.PerformLayout();
            this.gbMag.ResumeLayout(false);
            this.gbMag.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbFaculty;
        private System.Windows.Forms.ComboBox cbStudyBasis;
        private System.Windows.Forms.ComboBox cbStudyForm;
        private System.Windows.Forms.ComboBox cbLicenseProgram;
        private System.Windows.Forms.Button btnFillGrid;
        private System.Windows.Forms.DataGridView dgvAbitList;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox tbAbitsTop;
        private System.Windows.Forms.RadioButton rbAbitsTop;
        private System.Windows.Forms.RadioButton rbAbitsAll;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnToExcel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.PictureBox pictureBoxWhite;
        private System.Windows.Forms.Label labelWhite;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btn_GreenList;
        private System.Windows.Forms.PictureBox pictureBoxThistle;
        private System.Windows.Forms.Label labelThistle;
        private System.Windows.Forms.ToolTip ttextEntryView;
        private System.Windows.Forms.CheckBox cbZeroWave;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.RadioButton rbCrimea;
        private System.Windows.Forms.RadioButton rbForeign;
        private System.Windows.Forms.RadioButton rbCommon;
        private System.Windows.Forms.Button btnRePaint;
        private System.Windows.Forms.PictureBox pictureBoxBeige;
        private System.Windows.Forms.Label labelBeige;
        private System.Windows.Forms.Button btnRestoreOriginals;
        private System.Windows.Forms.GroupBox gb1kurs;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox tbDinamicWave;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.GroupBox gbMag;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox tbDinamicWaveMag;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnSetAllOrigins;
        private System.Windows.Forms.CheckBox chbOnlyWithOrigins;
        private System.Windows.Forms.CheckBox chbOnlyWithOriginsMag;
    }
}