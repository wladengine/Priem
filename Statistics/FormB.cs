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
    public partial class FormB : Form
    {
        public FormB()
        {
            InitializeComponent();
            Recalculate();
        }

        private void Recalculate()
        {
            using (PriemEntities context = new PriemEntities())
            {
                int kcp_b = context.Entry.Where(x => x.StudyLevel.LevelGroupId == 1 && x.StudyFormId == 1).Select(x => x.KCP ?? 0).DefaultIfEmpty(0).Sum();
                tbKCP_1kBudzh.Text = kcp_b.ToString();

                var enteredCnt = from Entered in context.extEntryView
                                 join Abit in context.Abiturient on Entered.AbiturientId equals Abit.Id
                                 where Abit.Entry.StudyLevel.LevelGroupId == 1 && Abit.Entry.StudyFormId == 1
                                 && Entered.Date <= dtpDate.Value
                                 select new
                                 {
                                     Abit.CompetitionId,
                                     Abit.Entry.StudyBasisId,
                                     Abit.Id
                                 };

                tbEnteredAbits_count.Text = enteredCnt.Count().ToString();

                tbEnteredAbitsCount_B.Text = enteredCnt.Where(x => x.StudyBasisId == 1).Count().ToString();
                tbEnteredAbitsCount_P.Text = enteredCnt.Where(x => x.StudyBasisId == 2).Count().ToString();

                tbEnteredAbitsCount_Cel_B.Text = enteredCnt.Where(x => x.StudyBasisId == 1 && x.CompetitionId == 6).Count().ToString();
                tbEnteredAbitsCount_Cel_P.Text = enteredCnt.Where(x => x.StudyBasisId == 2 && x.CompetitionId == 6).Count().ToString();

                tbEnteredAbitsCount_VK_B.Text = enteredCnt.Where(x => x.StudyBasisId == 1 && (x.CompetitionId == 2 || x.CompetitionId == 7)).Count().ToString();
                tbEnteredAbitsCount_VK_P.Text = enteredCnt.Where(x => x.StudyBasisId == 2 && (x.CompetitionId == 2 || x.CompetitionId == 7)).Count().ToString();

                var AbitIds = enteredCnt.Where(x => x.StudyBasisId == 1).Select(x => (Guid?)x.Id).Distinct().ToList();
                tbEnteredAbitsCount_Olymps_B.Text = context.Olympiads.Where(x => x.OlympValueId > 4 && AbitIds.Contains(x.AbiturientId)).Select(x => x.AbiturientId).Distinct().Count().ToString();
                AbitIds = enteredCnt.Where(x => x.StudyBasisId == 2).Select(x => (Guid?)x.Id).Distinct().ToList();
                tbEnteredAbitsCount_Olymps_P.Text = context.Olympiads.Where(x => x.OlympValueId > 4 && AbitIds.Contains(x.AbiturientId)).Select(x => x.AbiturientId).Distinct().Count().ToString();

                //AVG balls
                var EgeMarks =
                    (from EnteredAbits in context.extEntryView
                     join Abit in context.Abiturient on EnteredAbits.AbiturientId equals Abit.Id
                     join Mrk in context.Mark on Abit.Id equals Mrk.AbiturientId
                     where Abit.Entry.StudyFormId == 1 && Abit.Entry.StudyLevel.LevelGroupId == 1
                     && EnteredAbits.Date <= dtpDate.Value
                     select new { Value = (int)Mrk.Value, Abit.Entry.StudyBasisId });

                tbEnteredAbits_AvgEGE_B.Text = EgeMarks.Where(x => x.StudyBasisId == 1).Select(x => x.Value).Average().ToString();
                tbEnteredAbits_AvgEGE_P.Text = EgeMarks.Where(x => x.StudyBasisId == 2).Select(x => x.Value).DefaultIfEmpty(0).Average().ToString();
            }
        }

        private void dtpDate_ValueChanged(object sender, EventArgs e)
        {
            Recalculate();
        }
    }
}
