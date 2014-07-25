namespace Priem
{
    partial class CardLicenseProgram
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbName = new System.Windows.Forms.TextBox();
            this.tbNameEng = new System.Windows.Forms.TextBox();
            this.tbCode = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cbStudyLevel = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.epError)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(357, 133);
            // 
            // btnSaveChange
            // 
            this.btnSaveChange.Location = new System.Drawing.Point(12, 132);
            // 
            // btnSaveAsNew
            // 
            this.btnSaveAsNew.Location = new System.Drawing.Point(220, 132);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(50, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 25;
            this.label1.Text = "Название";
            // 
            // tbName
            // 
            this.tbName.Location = new System.Drawing.Point(113, 12);
            this.tbName.Name = "tbName";
            this.tbName.Size = new System.Drawing.Size(310, 20);
            this.tbName.TabIndex = 26;
            // 
            // tbNameEng
            // 
            this.tbNameEng.Location = new System.Drawing.Point(113, 38);
            this.tbNameEng.Name = "tbNameEng";
            this.tbNameEng.Size = new System.Drawing.Size(310, 20);
            this.tbNameEng.TabIndex = 27;
            // 
            // tbCode
            // 
            this.tbCode.Location = new System.Drawing.Point(113, 64);
            this.tbCode.Name = "tbCode";
            this.tbCode.Size = new System.Drawing.Size(115, 20);
            this.tbCode.TabIndex = 28;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 29;
            this.label2.Text = "Название англ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(81, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 13);
            this.label3.TabIndex = 30;
            this.label3.Text = "Код";
            // 
            // cbStudyLevel
            // 
            this.cbStudyLevel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbStudyLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyLevel.FormattingEnabled = true;
            this.cbStudyLevel.Location = new System.Drawing.Point(113, 90);
            this.cbStudyLevel.Name = "cbStudyLevel";
            this.cbStudyLevel.Size = new System.Drawing.Size(310, 21);
            this.cbStudyLevel.TabIndex = 31;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(56, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 32;
            this.label4.Text = "Уровень";
            // 
            // CardLicenseProgram
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(450, 167);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cbStudyLevel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbCode);
            this.Controls.Add(this.tbNameEng);
            this.Controls.Add(this.tbName);
            this.Controls.Add(this.label1);
            this.Name = "CardLicenseProgram";
            this.Text = "CardLicenseProgram";
            this.Controls.SetChildIndex(this.btnSaveChange, 0);
            this.Controls.SetChildIndex(this.btnClose, 0);
            this.Controls.SetChildIndex(this.btnSaveAsNew, 0);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.tbName, 0);
            this.Controls.SetChildIndex(this.tbNameEng, 0);
            this.Controls.SetChildIndex(this.tbCode, 0);
            this.Controls.SetChildIndex(this.label2, 0);
            this.Controls.SetChildIndex(this.label3, 0);
            this.Controls.SetChildIndex(this.cbStudyLevel, 0);
            this.Controls.SetChildIndex(this.label4, 0);
            ((System.ComponentModel.ISupportInitialize)(this.epError)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbName;
        private System.Windows.Forms.TextBox tbNameEng;
        private System.Windows.Forms.TextBox tbCode;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbStudyLevel;
        private System.Windows.Forms.Label label4;
    }
}