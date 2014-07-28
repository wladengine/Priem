namespace Priem
{
    partial class Querier
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
            this.btnPayDataEntryLoadCSV = new System.Windows.Forms.Button();
            this.btnImportExamSpecAspirant = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnPayDataEntryLoadCSV
            // 
            this.btnPayDataEntryLoadCSV.Location = new System.Drawing.Point(31, 25);
            this.btnPayDataEntryLoadCSV.Name = "btnPayDataEntryLoadCSV";
            this.btnPayDataEntryLoadCSV.Size = new System.Drawing.Size(132, 23);
            this.btnPayDataEntryLoadCSV.TabIndex = 0;
            this.btnPayDataEntryLoadCSV.Text = "PayDataEntry - load csv";
            this.btnPayDataEntryLoadCSV.UseVisualStyleBackColor = true;
            this.btnPayDataEntryLoadCSV.Click += new System.EventHandler(this.btnPayDataEntryLoadCSV_Click);
            // 
            // btnImportExamSpecAspirant
            // 
            this.btnImportExamSpecAspirant.Location = new System.Drawing.Point(31, 74);
            this.btnImportExamSpecAspirant.Name = "btnImportExamSpecAspirant";
            this.btnImportExamSpecAspirant.Size = new System.Drawing.Size(132, 23);
            this.btnImportExamSpecAspirant.TabIndex = 1;
            this.btnImportExamSpecAspirant.Text = "ImportExamSpecAspirant";
            this.btnImportExamSpecAspirant.UseVisualStyleBackColor = true;
            this.btnImportExamSpecAspirant.Click += new System.EventHandler(this.btnImportExamSpecAspirant_Click);
            // 
            // Querier
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(430, 319);
            this.Controls.Add(this.btnImportExamSpecAspirant);
            this.Controls.Add(this.btnPayDataEntryLoadCSV);
            this.Name = "Querier";
            this.Text = "Querier";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnPayDataEntryLoadCSV;
        private System.Windows.Forms.Button btnImportExamSpecAspirant;
    }
}