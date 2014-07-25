using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using BaseFormsLib;
using System.IO;
using System.Data.Objects;

namespace Priem
{
    public partial class VuzNameList : BaseList
    {
        public VuzNameList()
        {
            InitializeComponent();
            InitControls();
        }

        protected override void ExtraInit()
        {
            this.MdiParent = MainClass.mainform;
            _tableName = "ed.SchoolNames";
            _sQuery = "SELECT Id, Name FROM " + _tableName + " WHERE SchoolTypeId=4";
            Dgv = dgv;
        }

        protected override void OpenCard(string itemId)
        {
            var crd = new VuzNameCard(itemId);
            crd.ToUpdateList += UpdateDataGrid;
            crd.Show();
        }

        public override void UpdateDataGrid()
        {
            Dgv.DataSource = MainClass.Bdc.GetDataSet(_sQuery).Tables[0];
            Dgv.Columns["Id"].Visible = false;
        }

        protected override void Delete(string tableName, string id)
        {
            int iId = 0;
            int.TryParse(id, out iId);

            using (PriemEntities context = new PriemEntities())
            {
                context.SchoolNames_Delete(iId);
            }
        }

        private void btnFromList_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV files|*.csv";
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;
            int cntUpd = 0, cntIns = 0, cntFail = 0;
            string fName = ofd.FileName;
            using (PriemEntities context = new PriemEntities())
            using (StreamReader sr = new StreamReader(fName, Encoding.GetEncoding(1251)))
            {
                Dictionary<string, int> dicRegions = context.Region.Select(x => new { x.Id, x.Name }).ToList().ToDictionary(x => x.Name, y => y.Id);
                while (!sr.EndOfStream)
                {
                    string str = sr.ReadLine();
                    string[] splitted = str.Split(';');
                    if (splitted.Count() < 2)
                        continue;

                    string RegionName = splitted[0];
                    int RegionId = 0;
                    string Name = splitted[1];

                    if (dicRegions.ContainsKey(RegionName))
                        RegionId = dicRegions[RegionName];
                    if (RegionId != 0)
                    {
                        var nameIds = context.SchoolNames.Where(x => x.SchoolTypeId == 4 && x.RegionId == RegionId && x.Name == Name).Select(x => x.Id);
                        int cnt = nameIds.Count();
                        if (cnt > 0)
                        {
                            context.SchoolNames_Update(RegionId, 4, Name, nameIds.First());
                            cntUpd++;
                        }
                        else
                        {
                            context.SchoolNames_Insert(RegionId, 4, Name, new ObjectParameter("id", typeof(Guid)));
                            cntIns++;
                        }
                    }
                    else
                    {
                        cntFail++;
                    }
                }
            }
            MessageBox.Show("Done! CntIns=" + cntIns.ToString() + "; CntUpd=" + cntUpd.ToString() + "; CntFail=" + cntFail.ToString());
        }
    }
}
