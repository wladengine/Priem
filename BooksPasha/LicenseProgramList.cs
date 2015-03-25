using EducServLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Transactions;
using System.Windows.Forms;

namespace Priem
{
    public partial class LicenseProgramList : BookList
    {
        private int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
        }
        public LicenseProgramList()
        {
            InitializeComponent();
            InitControls();
        }

        protected override void ExtraInit()
        {
            Dgv = dgv;
            base.ExtraInit();
            _tableName = "ed.SP_LicenseProgram";
            _title = "Список направлений";
            
            using (PriemEntities context = new PriemEntities())
            {
                var data = context.StudyLevel.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyLevel, data, false, true);
            }
        }

        public override void InitHandlers()
        {
            cbStudyLevel.SelectedIndexChanged += cbStudyLevel_SelectedIndexChanged;
        }
        void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        public override void UpdateDataGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var data = context.SP_LicenseProgram.Where(x => StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId.Value : true).Select(x => new { x.Id, x.Code, x.Name }).OrderBy(x => x.Code).ToArray();
                dgv.DataSource = Converter.ConvertToDataTable(data);
            }

            Dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            lblCount.Text = "Всего: " + dgv.Rows.Count.ToString();
            SetVisibleColumnsAndNameColumns();
            Dgv.Columns["Code"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            Dgv.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        protected override void SetVisibleColumnsAndNameColumns()
        {
            base.SetVisibleColumnsAndNameColumns();
            SetVisibleColumnsAndNameColumns("Code", "Шифр");
        }

        protected override void OpenCard(string itemId)
        {
            var crd = new CardLicenseProgram(itemId);
            crd.ToUpdateList += UpdateDataGrid;
            crd.Show();
        }

        protected override void DeleteSelectedRows(string sId)
        {
            using (PriemEntities context = new PriemEntities())
            using (TransactionScope tran = new TransactionScope())
            {
                int iId = 0;
                int.TryParse(sId, out iId);
                context.SP_LicenseProgram_Delete(iId);

                string query = "DELETE FROM SP_LicenseProgram WHERE Id=@Id";
                MainClass.BdcOnlineReadWrite.ExecuteQuery(query, new SortedList<string, object>() { { "@Id", iId } });

                tran.Complete();
            }
        }
    }
}
