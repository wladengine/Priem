using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using BaseFormsLib;
using BDClassLib;

namespace Priem
{
    public partial class DocCard : BaseForm
    {
        private DocsClass _docs;
        private int _personBarc;
        private int? _abitBarc;
        private bool _upd;

        public DocCard(int perBarcode, int? abitBarcode, bool upd)
        {
            InitializeComponent();
            _personBarc = perBarcode;
            _abitBarc = abitBarcode;
            _docs = new DocsClass(_personBarc, _abitBarc);
            _upd = upd;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            InitControls();
        }

        private void InitControls()
        {
            InitFocusHandlers();

            this.CenterToParent();

            dgvFiles.DataSource = _docs.UpdateFilesTable();
            if (dgvFiles.Rows.Count > 0)
            {
                foreach (DataGridViewColumn clm in dgvFiles.Columns)
                    clm.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                
                if (!dgvFiles.Columns.Contains("Открыть"))
                {
                    DataGridViewCheckBoxCell cl = new DataGridViewCheckBoxCell();
                    cl.TrueValue = true;
                    cl.FalseValue = false;

                    DataGridViewCheckBoxColumn clm = new DataGridViewCheckBoxColumn();
                    clm.CellTemplate = cl;
                    clm.Name = "Открыть";
                    dgvFiles.Columns.Add(clm);
                    dgvFiles.Columns["Открыть"].DisplayIndex = 0;
                    dgvFiles.Columns["Открыть"].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader; 
                }
                if (dgvFiles.Columns.Contains("Id"))
                    dgvFiles.Columns["Id"].Visible = false;
                if (dgvFiles.Columns.Contains("FileExtention"))
                    dgvFiles.Columns["FileExtention"].Visible = false;
                dgvFiles.Columns["FileName"].HeaderText = "Файл";
                dgvFiles.Columns["FileName"].ReadOnly = true;
                
                dgvFiles.Columns["Comment"].HeaderText = "Комментарий";
                dgvFiles.Columns["Comment"].ReadOnly = true;

                dgvFiles.Columns["FileTypeName"].HeaderText = "Тип файла";
                dgvFiles.Columns["FileTypeName"].ReadOnly = true;

            } 
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();
            lstFiles = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow rw in dgvFiles.Rows)
            {
                DataGridViewCheckBoxCell cell = rw.Cells["Открыть"] as DataGridViewCheckBoxCell;
                if (cell.Value == cell.TrueValue)
                {
                    if (dgvFiles.Columns.Contains("FileName"))
                    {
                        string fileName = rw.Cells["FileName"].Value.ToString();
                        KeyValuePair<string, string> file = new KeyValuePair<string, string>(rw.Cells["Id"].Value.ToString(), fileName);
                        lstFiles.Add(file);
                    }
                }
            }
            _docs.OpenFile(lstFiles);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DocCard_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_upd)
                if (_docs != null)
                {
                    _docs.BDCInet.ExecuteQuery(string.Format("UPDATE Person SET DateReviewDocs = '{0}' WHERE Person.Barcode = {1}", DateTime.Now.ToString(), _personBarc));
                    if(_abitBarc != null)
                        _docs.BDCInet.ExecuteQuery(string.Format("UPDATE Application SET DateReviewDocs = '{0}' WHERE Application.Barcode = {1}", DateTime.Now.ToString(), _abitBarc));
                
                    _docs.CloseDB();
                }
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow rw in dgvFiles.Rows)
            {
                DataGridViewCheckBoxCell cell = rw.Cells["Открыть"] as DataGridViewCheckBoxCell;
                cell.Value = cell.TrueValue;
            }
        }

        private void btnCheckNone_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow rw in dgvFiles.Rows)
            {
                DataGridViewCheckBoxCell cell = rw.Cells["Открыть"] as DataGridViewCheckBoxCell;
                cell.Value = cell.FalseValue;
            }
        }
    }
}
