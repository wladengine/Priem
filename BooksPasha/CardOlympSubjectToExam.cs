using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;

namespace Priem
{
    public partial class CardOlympSubjectToExam : BookCardInt
    {
        private int OlympSubjectId
        {
            get;
            set;
        }
        private int? ExamId
        {
            get { return ComboServ.GetComboIdInt(cbExam); }
            set { ComboServ.SetComboId(cbExam, value); }
        }

        public CardOlympSubjectToExam(string id)
        {
            InitializeComponent();
            _Id = id;
            InitControls();
        }

        protected override void ExtraInit()
        {
            FillCombo();
            base.ExtraInit();
        }

        protected override void FillCard()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = context.OlympSubjectToExam.Where(x => x.Id == IntId).FirstOrDefault();
                if (ent != null)
                {
                    OlympSubjectId = ent.OlympSubjectId;
                    UpdateOlympSubject();
                    ExamId = ent.ExamId;
                }
            }
        }

        private void UpdateOlympSubject()
        {
            using (PriemEntities context = new PriemEntities())
            {
                tbOlympSubject.Text = context.OlympSubject.Where(x => x.Id == OlympSubjectId).Select(x => x.Name).FirstOrDefault();
            }
        }

        private void FillCombo()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.extExamInEntry.Where(x => x.StudyLevelGroupId == MainClass.studyLevelGroupId)
                    .Select(x => new { x.ExamId, x.ExamName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.ExamId.ToString(), x.ExamName)).ToList();

                ComboServ.FillCombo(cbExam, src, false, false);
            }
        }

        public CardOlympSubjectToExam(int olympSubjectId)
        {
            InitializeComponent();

            OlympSubjectId = olympSubjectId;
            UpdateOlympSubject();
            InitControls();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveClick();
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.OlympSubjectToExam_Insert(OlympSubjectId, ExamId, idParam);
        }

        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.OlympSubjectToExam_Update(OlympSubjectId, ExamId, id);
        }
    }
}
