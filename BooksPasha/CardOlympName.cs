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
    public partial class CardOlympName : BookCardInt
    {
        private string OlympName
        {
            get { return tbName.Text.Trim(); }
            set { tbName.Text = value; }
        }
        private int? OlympNumber
        {
            get 
            {
                int iRet;
                if (string.IsNullOrEmpty(tbNumber.Text) || !int.TryParse(tbNumber.Text, out iRet))
                    return null;
                else
                    return iRet;
            }
            set
            {
                tbNumber.Text = value.HasValue ? value.Value.ToString() : "";
            }
        }

        public CardOlympName(string id) : base(id)
        {
            InitializeComponent();
            _tableName = "ed.OlympName";
            _title = "Название олимпиады";
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

                var ent = context.OlympName.Where(x => x.Id == IntId).First();
                OlympName = ent.Name;
                OlympNumber = ent.Number;
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.OlympName_Insert(OlympName, OlympNumber, idParam);
            string query = "INSERT INTO OlympName(Id, Name, Number) VALUES (@Id, @Name, @Number)";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", idParam.Value);
            slParams.AddVal("@Name", OlympName);
            slParams.AddVal("@Number", OlympNumber);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }

        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.OlympName_Update(OlympName, OlympNumber, id);
            string query = "UPDATE OlympName SET [Name]=@Name, Number=@Number WHERE Id=@id";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", id);
            slParams.AddVal("@Name", OlympName);
            slParams.AddVal("@Number", OlympNumber);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }
    }
}
