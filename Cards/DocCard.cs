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

        public DocCard(int perBarcode, int? abitBarcode)
        {
            InitializeComponent();
            _personBarc = perBarcode;
            _abitBarc = abitBarcode;
            _docs = new DocsClass(_personBarc, _abitBarc);

            InitControls();
        }

        private void InitControls()
        {
            InitFocusHandlers();

            this.CenterToParent();

            dgvFiles.DataSource = _docs.UpdateFilesTable();
            if (dgvFiles.Rows.Count > 0)
            {
                dgvFiles.ReadOnly = false;
                foreach (DataGridViewColumn clm in dgvFiles.Columns)
                {
                    clm.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }
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
                
                foreach (DataGridViewRow rw in dgvFiles.Rows)
                {
                    string filename = rw.Cells["FileName"].Value.ToString();
                    string fileext = rw.Cells["FileExtention"].Value.ToString();
                    if (!String.IsNullOrEmpty(fileext))
                    {
                        if (filename.EndsWith(fileext))
                            filename = filename.Substring(0, filename.Length - fileext.Length);
                    }
                    filename += " (" + rw.Cells["LoadDate"].Value.ToString() + ")";
                    filename += fileext;
                    rw.Cells["FileName"].Value = filename;
                    rw.ReadOnly = false;
                    rw.Cells["Открыть"].ReadOnly = false;
                    rw.Cells["FileName"].ReadOnly = true;
                    rw.Cells["Comment"].ReadOnly = true;
                }
                dgvFiles.Columns["LoadDate"].Visible = false;

            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow rw in dgvFiles.Rows)
            {
                DataGridViewCheckBoxCell cell = rw.Cells["Открыть"] as DataGridViewCheckBoxCell;
                if (cell.Value == cell.TrueValue)
                {
                    if (dgvFiles.Columns.Contains("Файл"))
                    {
                        string fileName = rw.Cells["Файл"].Value.ToString();
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
            if (_docs != null)
            {
                _docs.BDCInet.ExecuteQuery(string.Format("UPDATE Person SET DateReviewDocs = '{0}' WHERE Person.Barcode = {1}", DateTime.Now.ToString(), _personBarc));
                if(_abitBarc != null)
                    _docs.BDCInet.ExecuteQuery(string.Format("UPDATE Application SET DateReviewDocs = '{0}' WHERE Application.Barcode = {1}", DateTime.Now.ToString(), _abitBarc));
                
                _docs.CloseDB();
            }
        }
    }
}
