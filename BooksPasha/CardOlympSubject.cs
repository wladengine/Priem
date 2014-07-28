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

namespace Priem
{
    public partial class CardOlympSubject : BookCardInt
    {
        private string OlympName
        {
            get { return tbName.Text.Trim(); }
            set { tbName.Text = value; }
        }
        private string OlympNameDative
        {
            get { return tbNameDative.Text.Trim(); }
            set { tbNameDative.Text = value; }
        }
        
        public CardOlympSubject(string id)
            : base(id)
        {
            InitializeComponent();
            _tableName = "ed.OlympSubject";
            _title = "Предмет олимпиады";
            InitControls();
        }

        protected override void FillCard()
        {
            using (PriemEntities context = new PriemEntities())
            {
                int IntId = 0;
                int.TryParse(_Id, out IntId);
                if (IntId == 0)
                    return;

                var ent = context.OlympSubject.Where(x => x.Id == IntId).First();
                OlympName = ent.Name;
                OlympNameDative = ent.NameDative;
                UpdateGrid();
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.OlympSubject_Insert(OlympName, OlympNameDative, idParam);
            string query = "INSERT INTO OlympSubject (Id, [Name]) values (@Id, @Name)";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", idParam.Value);
            slParams.AddVal("@Name", OlympName);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }

        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.OlympSubject_Update(OlympName, OlympNameDative, id);
            string query = "UPDATE OlympSubject SET [Name]=@Name WHERE Id=@id";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", id);
            slParams.AddVal("@Name", OlympName);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }


        private void UpdateGrid()
        {
            if (!IntId.HasValue)
                return;

            using (PriemEntities context = new PriemEntities())
            {
                var data = context.OlympSubjectToExam.Where(x => x.OlympSubjectId == IntId).Select(x => new { x.Id, x.Exam.ExamName.Name }).ToList();
                dgvOlympSubjectToExam.DataSource = data;
                dgvOlympSubjectToExam.Columns["Id"].Visible = false;
                dgvOlympSubjectToExam.Columns["Name"].HeaderText = "Предмет";
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!IntId.HasValue)
                return;

            OpenCardOlympSubjectToExam(null);
        }

        private void OpenCardOlympSubjectToExam(string sId)
        {
            if (!IntId.HasValue)
                return;

            CardOlympSubjectToExam crd;

            if (string.IsNullOrEmpty(sId))
                crd = new CardOlympSubjectToExam(IntId.Value);
            else
                crd = new CardOlympSubjectToExam(sId);

            crd.ToUpdateList += UpdateGrid;
            crd.Show();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (dgvOlympSubjectToExam.SelectedCells.Count == 0)
                return;

            int rwInd = dgvOlympSubjectToExam.SelectedCells[0].RowIndex;
            OpenCardOlympSubjectToExam(dgvOlympSubjectToExam["Id", rwInd].Value.ToString());
        }
        private void dgvOlympSubjectToExam_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            int rwInd = e.RowIndex;
            OpenCardOlympSubjectToExam(dgvOlympSubjectToExam["Id", rwInd].Value.ToString());
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvOlympSubjectToExam.SelectedCells.Count == 0)
                return;

            int rwInd = dgvOlympSubjectToExam.SelectedCells[0].RowIndex;
            int iId = (int)dgvOlympSubjectToExam["Id", rwInd].Value;

            using (PriemEntities context = new PriemEntities())
            {
                context.OlympSubjectToExam_Delete(iId);
            }

            UpdateGrid();
        }
    }
}
