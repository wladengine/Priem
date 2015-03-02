namespace Priem
{
    partial class CardInnerEntryInEntryInCompetitionInInet
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
            this.tbLicenseProgramName = new System.Windows.Forms.TextBox();
            this.dgvObrazProgramInEntryList = new System.Windows.Forms.DataGridView();
            this.lblHasAlpabetSort = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvObrazProgramInEntryList)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Направление";
            // 
            // tbLicenseProgramName
            // 
            this.tbLicenseProgramName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbLicenseProgramName.Location = new System.Drawing.Point(93, 12);
            this.tbLicenseProgramName.Name = "tbLicenseProgramName";
            this.tbLicenseProgramName.ReadOnly = true;
            this.tbLicenseProgramName.Size = new System.Drawing.Size(404, 20);
            this.tbLicenseProgramName.TabIndex = 1;
            // 
            // dgvObrazProgramInEntryList
            // 
            this.dgvObrazProgramInEntryList.AllowUserToAddRows = false;
            this.dgvObrazProgramInEntryList.AllowUserToDeleteRows = false;
            this.dgvObrazProgramInEntryList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvObrazProgramInEntryList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvObrazProgramInEntryList.Location = new System.Drawing.Point(12, 38);
            this.dgvObrazProgramInEntryList.Name = "dgvObrazProgramInEntryList";
            this.dgvObrazProgramInEntryList.ReadOnly = true;
            this.dgvObrazProgramInEntryList.Size = new System.Drawing.Size(646, 352);
            this.dgvObrazProgramInEntryList.TabIndex = 2;
            // 
            // lblHasAlpabetSort
            // 
            this.lblHasAlpabetSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblHasAlpabetSort.ForeColor = System.Drawing.Color.DarkRed;
            this.lblHasAlpabetSort.Location = new System.Drawing.Point(503, 5);
            this.lblHasAlpabetSort.Name = "lblHasAlpabetSort";
            this.lblHasAlpabetSort.Size = new System.Drawing.Size(155, 30);
            this.lblHasAlpabetSort.TabIndex = 3;
            this.lblHasAlpabetSort.Text = "Внимание! Приоритеты отсортированы по алфавиту!";
            this.lblHasAlpabetSort.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblHasAlpabetSort.Visible = false;
            // 
            // CardInnerEntryInEntryInCompetitionInInet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(670, 402);
            this.Controls.Add(this.lblHasAlpabetSort);
            this.Controls.Add(this.dgvObrazProgramInEntryList);
            this.Controls.Add(this.tbLicenseProgramName);
            this.Controls.Add(this.label1);
            this.Name = "CardInnerEntryInEntryInCompetitionInInet";
            this.Text = "Приоритеты образовательных программ и профилей в конкурсе";
            ((System.ComponentModel.ISupportInitialize)(this.dgvObrazProgramInEntryList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbLicenseProgramName;
        private System.Windows.Forms.DataGridView dgvObrazProgramInEntryList;
        private System.Windows.Forms.Label lblHasAlpabetSort;
    }
}