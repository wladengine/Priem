namespace Priem.BooksPasha
{
    partial class ParametersForm
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
            this.btnSave = new System.Windows.Forms.Button();
            this.chb1kCheckProtocolsEnabled = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbPriemYear = new System.Windows.Forms.TextBox();
            this.chbMagCheckProtocolsEnabled = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(12, 340);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // chb1kCheckProtocolsEnabled
            // 
            this.chb1kCheckProtocolsEnabled.AutoSize = true;
            this.chb1kCheckProtocolsEnabled.Location = new System.Drawing.Point(6, 19);
            this.chb1kCheckProtocolsEnabled.Name = "chb1kCheckProtocolsEnabled";
            this.chb1kCheckProtocolsEnabled.Size = new System.Drawing.Size(187, 17);
            this.chb1kCheckProtocolsEnabled.TabIndex = 1;
            this.chb1kCheckProtocolsEnabled.Text = "Проверять при запуске (1 курс)";
            this.chb1kCheckProtocolsEnabled.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Текущий год приёма";
            // 
            // tbPriemYear
            // 
            this.tbPriemYear.Location = new System.Drawing.Point(12, 25);
            this.tbPriemYear.Name = "tbPriemYear";
            this.tbPriemYear.Size = new System.Drawing.Size(103, 20);
            this.tbPriemYear.TabIndex = 3;
            // 
            // chbMagCheckProtocolsEnabled
            // 
            this.chbMagCheckProtocolsEnabled.AutoSize = true;
            this.chbMagCheckProtocolsEnabled.Location = new System.Drawing.Point(6, 42);
            this.chbMagCheckProtocolsEnabled.Name = "chbMagCheckProtocolsEnabled";
            this.chbMagCheckProtocolsEnabled.Size = new System.Drawing.Size(225, 17);
            this.chbMagCheckProtocolsEnabled.TabIndex = 4;
            this.chbMagCheckProtocolsEnabled.Text = "Проверять при запуске (магистратура)";
            this.chbMagCheckProtocolsEnabled.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chb1kCheckProtocolsEnabled);
            this.groupBox1.Controls.Add(this.chbMagCheckProtocolsEnabled);
            this.groupBox1.Location = new System.Drawing.Point(12, 51);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(237, 69);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Проверка свежих протоколов о допуске";
            // 
            // ParametersForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(712, 375);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.tbPriemYear);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSave);
            this.Name = "ParametersForm";
            this.Text = "Параметры \"Приёма\"";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.CheckBox chb1kCheckProtocolsEnabled;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbPriemYear;
        private System.Windows.Forms.CheckBox chbMagCheckProtocolsEnabled;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}