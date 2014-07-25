using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Priem
{
    public partial class CardExaminerInExamsVed : Form
    {
        private List<string> lstExaminers
        {
            get { return mslExaminer.SelectedList.ToList(); }
            set { mslExaminer.SelectedList = value; }
        }
        private Guid ExamsVedId;

        public CardExaminerInExamsVed(Guid _ExamsVedId)
        {
            InitializeComponent();
            ExamsVedId = _ExamsVedId;
            this.MdiParent = MainClass.mainform;

            InitExaminerList();
            FillCard();
        }

        private void InitExaminerList()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.ExaminerAccount
                    .Select(x => new {x.AccountName, x.DisplayName })
                    .ToList()
                    .OrderBy(x => x.DisplayName)
                    .ToDictionary(x => x.AccountName, y => y.DisplayName + " (" + y.AccountName + ")");
                mslExaminer.InitSource(src);
            }
        }

        private void FillCard()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var vedInfo = context.extExamsVed.Where(x => x.Id == ExamsVedId).FirstOrDefault();
                if (vedInfo != null)
                {
                    tbExamsVed.Text = vedInfo.ExamName + " " + vedInfo.Date;
                }
                var examiners = context.ExaminerInExamsVed.Where(x => x.ExamsVedId == ExamsVedId).Select(x => x.ExaminerAccount).ToList();
                if (examiners.Count > 0)
                {
                    lstExaminers = examiners;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Сохранить список экзаменаторов для ведомости?", "Сохранение", MessageBoxButtons.YesNoCancel) == System.Windows.Forms.DialogResult.Yes)
            {
                SaveCard();
                this.Close();
            }
        }

        private void SaveCard()
        {
            using (PriemEntities context = new PriemEntities())
            {
                context.ExamsVed_DeleteAllExaminerAccountInVed(ExamsVedId);

                foreach (string accouns in lstExaminers)
                {
                    context.ExamsVed_SetExaminerAccount(accouns, ExamsVedId);
                }
            }
        }
    }
}
