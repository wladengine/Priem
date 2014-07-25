namespace Priem
{
    partial class CardProfileInObrazProgramInEntry
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
            this.label2 = new System.Windows.Forms.Label();
            this.tbLicenseProgramName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbProfile = new System.Windows.Forms.ComboBox();
            this.tbObrazProgramName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbKCP = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(34, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Направление";
            // 
            // tbLicenseProgramName
            // 
            this.tbLicenseProgramName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbLicenseProgramName.Location = new System.Drawing.Point(115, 12);
            this.tbLicenseProgramName.Name = "tbLicenseProgramName";
            this.tbLicenseProgramName.ReadOnly = true;
            this.tbLicenseProgramName.Size = new System.Drawing.Size(302, 20);
            this.tbLicenseProgramName.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Обр. программа";
            // 
            // cbProfile
            // 
            this.cbProfile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbProfile.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbProfile.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbProfile.FormattingEnabled = true;
            this.cbProfile.Location = new System.Drawing.Point(115, 64);
            this.cbProfile.Name = "cbProfile";
            this.cbProfile.Size = new System.Drawing.Size(302, 21);
            this.cbProfile.TabIndex = 1;
            // 
            // tbObrazProgramName
            // 
            this.tbObrazProgramName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbObrazProgramName.Location = new System.Drawing.Point(115, 38);
            this.tbObrazProgramName.Name = "tbObrazProgramName";
            this.tbObrazProgramName.ReadOnly = true;
            this.tbObrazProgramName.Size = new System.Drawing.Size(302, 20);
            this.tbObrazProgramName.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(56, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Профиль";
            // 
            // tbKCP
            // 
            this.tbKCP.Location = new System.Drawing.Point(115, 91);
            this.tbKCP.Name = "tbKCP";
            this.tbKCP.Size = new System.Drawing.Size(100, 20);
            this.tbKCP.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(79, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "КЦП";
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSave.Location = new System.Drawing.Point(12, 132);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // CardProfileInObrazProgramInEntry
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 167);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.tbKCP);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbObrazProgramName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbLicenseProgramName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbProfile);
            this.Name = "CardProfileInObrazProgramInEntry";
            this.Text = "Профиль в образовательной программе в году";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbLicenseProgramName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbProfile;
        private System.Windows.Forms.TextBox tbObrazProgramName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbKCP;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnSave;
    }
}