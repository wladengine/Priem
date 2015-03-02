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
    public partial class CardInnerEntryInEntryInCompetitionInInet : Form
    {
        private ShortCompetition comp;
        public CardInnerEntryInEntryInCompetitionInInet(ShortCompetition _comp)
        {
            this.MdiParent = MainClass.mainform;
            InitializeComponent();
            comp = _comp;
            tbLicenseProgramName.Text = _comp.LicenseProgramName;
            InitGrid();
        }

        private void InitGrid()
        {
            var src =
                comp.lstObrazProgramsInEntry.OrderBy(x => x.InnerEntryInEntryPriority)
                .Select(x => new
                {
                    x.Id,
                    x.InnerEntryInEntryPriority,
                    x.ObrazProgramName,
                    x.ProfileName
                }).ToList();

            dgvObrazProgramInEntryList.DataSource = src;
            dgvObrazProgramInEntryList.Columns["Id"].Visible = false;

            dgvObrazProgramInEntryList.Columns["InnerEntryInEntryPriority"].HeaderText = "Приоритет";
            dgvObrazProgramInEntryList.Columns["ObrazProgramName"].HeaderText = "Обр. программа";
            dgvObrazProgramInEntryList.Columns["ObrazProgramName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvObrazProgramInEntryList.Columns["ProfileName"].HeaderText = "Профиль";
            dgvObrazProgramInEntryList.Columns["ProfileName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //проверка на сортировку
            int i = 0;
            var OPS = comp.lstObrazProgramsInEntry.OrderBy(x => x.ObrazProgramName).ThenBy(x => x.ProfileName)
                .Select(x => new
                {
                    x.Id,
                    x.InnerEntryInEntryPriority,
                    x.ObrazProgramName,
                    x.ProfileName
                }).ToList();
            bool bIsSortedByAlpabet = true;
            foreach (var xOP in OPS)
            {
                if (src[i].ObrazProgramName != xOP.ObrazProgramName || src[i].ProfileName != xOP.ProfileName)
                    bIsSortedByAlpabet = false;
            }

            lblHasAlpabetSort.Visible = bIsSortedByAlpabet;
        }
    }
}
