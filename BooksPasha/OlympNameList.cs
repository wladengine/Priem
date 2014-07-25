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
    public partial class OlympNameList : BookList
    {
        public OlympNameList()
        {
            InitializeComponent();

            Dgv = dgvOlymps;
            _tableName = "ed.OlympName";
            _title = "Список названий олимпиад";

            InitControls();
        }

        protected override void OpenCard(string id, BaseFormEx formOwner, int? index)
        {
            var crd = new CardOlympName(id);
            crd.ToUpdateList += UpdateDataGrid;
            crd.Show();
        }

        protected override void GetSource()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var query = context.OlympName.Select(x => new { x.Id, x.Name, x.Number });

                Dgv.DataSource = query.ToList();
                SetVisibleColumnsAndNameColumns();
            }
        }

        protected override void SetVisibleColumnsAndNameColumns()
        {
            Dgv.AutoGenerateColumns = false;

            foreach (DataGridViewColumn col in Dgv.Columns)
            {
                col.Visible = false;
            }

            SetVisibleColumnsAndNameColumnsOrdered("Name", "Название", 1);
            SetVisibleColumnsAndNameColumnsOrdered("Number", "Номер", 2);

            Dgv.Columns["Name"].Width = 250;
            Dgv.Columns["Number"].Width = 50;
        }

        protected override void DeleteSelectedRows(string sId)
        {
            if (!MainClass.IsEntryChanger())
                return;

            using (PriemEntities context = new PriemEntities())
            {
                int IntId = 0;
                int.TryParse(sId, out IntId);
                context.OlympName_Delete(IntId);

                string query = "DELETE FROM OlympName WHERE Id=@Id";
                SortedList<string, object> slParams = new SortedList<string, object>();
                slParams.Add("@Id", IntId);

                MainClass.BdcOnlineReadWrite.ExecuteQuery(query);
            }
        }
    }
}
