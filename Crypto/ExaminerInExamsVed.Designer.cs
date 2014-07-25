namespace Priem
{
    partial class CardExaminerInExamsVed
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
            this.mslExaminer = new BaseFormsLib.MultySelectList();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.tbExamsVed = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // mslExaminer
            // 
            this.mslExaminer.Location = new System.Drawing.Point(12, 51);
            this.mslExaminer.Name = "mslExaminer";
            this.mslExaminer.SelectedList = null;
            this.mslExaminer.Size = new System.Drawing.Size(532, 247);
            this.mslExaminer.TabIndex = 0;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(12, 304);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(469, 304);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            // 
            // tbExamsVed
            // 
            this.tbExamsVed.Location = new System.Drawing.Point(12, 25);
            this.tbExamsVed.Name = "tbExamsVed";
            this.tbExamsVed.ReadOnly = true;
            this.tbExamsVed.Size = new System.Drawing.Size(532, 20);
            this.tbExamsVed.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Ведомость";
            // 
            // CardExaminerInExamsVed
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(556, 339);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbExamsVed);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.mslExaminer);
            this.Name = "CardExaminerInExamsVed";
            this.Text = "Экзаменаторы для экзаменационной ведомости";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private BaseFormsLib.MultySelectList mslExaminer;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TextBox tbExamsVed;
        private System.Windows.Forms.Label label1;
    }
}