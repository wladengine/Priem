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
        private int? ExamId
        {
            get { return ComboServ.GetComboIdInt(cbExam); }
            set { ComboServ.SetComboId(cbExam, value); }
        }
        public CardOlympSubject(string id)
            : base(id)
        {
            InitializeComponent();
            _tableName = "ed.OlympSubject";
            _title = "Предмет олимпиады";
            InitControls();
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.extExamInEntry.Where(x => x.StudyLevelGroupId == MainClass.studyLevelGroupId).Select(x => new { x.ExamId, x.ExamName })
                    .Distinct().ToList().Select(x => new KeyValuePair<string, string>(x.ExamId.ToString(), x.ExamName)).ToList();

                ComboServ.FillCombo(cbExam, src, true, false);
            }
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
                ExamId = ent.ExamId;
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.OlympSubject_Insert(OlympName, OlympNameDative, ExamId, idParam);
            string query = "INSERT INTO OlympSubject (Id, [Name]) values (@Id, @Name)";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", idParam.Value);
            slParams.AddVal("@Name", OlympName);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }

        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.OlympSubject_Update(OlympName, OlympNameDative, ExamId, id);
            string query = "UPDATE OlympSubject SET [Name]=@Name WHERE Id=@id";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", id);
            slParams.AddVal("@Name", OlympName);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }
    }
}
