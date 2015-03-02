using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BaseFormsLib;
using System.Transactions;

namespace Priem
{
    public partial class ObrazProgramList : BookList
    {
        private int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
        }
        private int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
        }
        public ObrazProgramList()
        {
            InitializeComponent();
            InitControls();
        }

        protected override void ExtraInit()
        {
            Dgv = dgv;
            base.ExtraInit();
            this.MdiParent = MainClass.mainform;
            _tableName = "ed.SP_ObrazProgram";
            _title = "Список образовательных программ";
            
            using (PriemEntities context = new PriemEntities())
            {
                var data = context.SP_Faculty.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbFaculty, data, false, true);

                data = context.StudyLevel.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyLevel, data, false, true);
            }
        }

        public override void InitHandlers()
        {
            base.InitHandlers();
            cbFaculty.SelectedIndexChanged += cbFaculty_SelectedIndexChanged;

            cbStudyLevel.SelectedIndexChanged += cbStudyLevel_SelectedIndexChanged;
        }

        void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }
        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        public override void UpdateDataGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var data = context.SP_ObrazProgram.Where(x => (FacultyId.HasValue ? x.FacultyId == FacultyId.Value : true) && (StudyLevelId.HasValue ? x.SP_LicenseProgram.StudyLevelId == StudyLevelId : true))
                    .Select(x => new { x.Id, Number = x.SP_LicenseProgram.StudyLevel.Acronym + "." + x.Number + "." + MainClass.sPriemYear, x.Name }).OrderBy(x => x.Name).ToArray();
                dgv.DataSource = Util.ConvertToDataTable(data);
            }
            lblCount.Text = "Всего: " + dgv.Rows.Count.ToString();
            SetVisibleColumnsAndNameColumns();
            Dgv.Columns["Number"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            Dgv.Columns["Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            Dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        protected override void SetVisibleColumnsAndNameColumns()
        {
            base.SetVisibleColumnsAndNameColumns();
            SetVisibleColumnsAndNameColumns("Number", "Номер");
        }

        protected override void OpenCard(string itemId)
        {
            var crd = new CardObrazProgram(itemId, FacultyId, StudyLevelId);
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
                context.SP_ObrazProgram_Delete(iId);

                string query = "DELETE FROM SP_ObrazProgram WHERE Id=@Id";
                MainClass.BdcOnlineReadWrite.ExecuteQuery(query, new SortedList<string, object>() { { "@Id", iId } });

                tran.Complete();
            }
        }
    }
}
