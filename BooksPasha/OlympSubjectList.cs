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
    public partial class OlympSubjectList : BookList
    {
        public OlympSubjectList()
        {
            InitializeComponent();

            Dgv = dgvData;
            _tableName = "ed.OlympSubject";
            _title = "Список предметов олимпиад";

            InitControls();
        }

        protected override void OpenCard(string id, BaseFormEx formOwner, int? index)
        {
            var crd = new CardOlympSubject(id);
            crd.ToUpdateList += UpdateDataGrid;
            crd.Show();
        }

        protected override void GetSource()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var query = context.OlympSubject.Select(x => new { x.Id, x.Name, x.NameDative });

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
            SetVisibleColumnsAndNameColumnsOrdered("NameDative", "В дат. падеже", 2);

            Dgv.Columns["Name"].Width = 250;
            Dgv.Columns["NameDative"].Width = 50;
        }

        protected override void DeleteSelectedRows(string sId)
        {
            if (!MainClass.IsEntryChanger())
                return;

            using (PriemEntities context = new PriemEntities())
            {
                int IntId = 0;
                int.TryParse(sId, out IntId);
                context.OlympSubject_Delete(IntId);

                string query = "DELETE FROM OlympSubject WHERE Id=@Id";
                SortedList<string, object> slParams = new SortedList<string, object>();
                slParams.Add("@Id", IntId);

                MainClass.BdcOnlineReadWrite.ExecuteQuery(query);
            }
        }
    }
}
