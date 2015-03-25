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
    public partial class CardLicenseProgram : BookCardInt
    {
        private string EntityName
        {
            get { return tbName.Text.Trim(); }
            set { tbName.Text = value; }
        }
        private string EntityNameEng
        {
            get { return tbNameEng.Text.Trim(); }
            set { tbNameEng.Text = value; }
        }
        private string Code
        {
            get { return tbCode.Text.Trim(); }
            set { tbCode.Text = value; }
        }
        private int StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel).Value; }
            set { ComboServ.SetComboId(cbStudyLevel, value); }
        }
        private int AggregateGroupId
        {
            get { return ComboServ.GetComboIdInt(cbAggregateGroup).Value; }
            set { ComboServ.SetComboId(cbAggregateGroup, value); }
        }
        private string GSGUCode
        {
            get { return tbGSGUCode.Text.Trim(); }
            set { tbGSGUCode.Text = value; }
        }
        public CardLicenseProgram(string id) : base(id)
        {
            InitializeComponent();
            InitControls();
        }

        protected override void ExtraInit()
        {
            this.MdiParent = MainClass.mainform;
            _title = "Направление";
            _tableName = "ed.SP_LicenseProgram";

            using (PriemEntities context = new PriemEntities())
            {
                var src = context.StudyLevel.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyLevel, src, false, false);

                src = context.SP_AggregateGroup.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbAggregateGroup, src, false, false);
            }
        }

        protected override void FillCard()
        {
            if (string.IsNullOrEmpty(_Id))
                return;
            else
            {
                using (PriemEntities context = new PriemEntities())
                {
                    var LP = context.SP_LicenseProgram.Where(x => x.Id == IntId).FirstOrDefault();
                    if (LP == null)
                        WinFormsServ.Error("Не найдена запись в базе!");
                    else
                    {
                        EntityName = LP.Name;
                        EntityNameEng = LP.NameEng;
                        Code = LP.Code;
                        StudyLevelId = LP.StudyLevelId;
                        GSGUCode = LP.GSGUCode;
                        AggregateGroupId = LP.AggregateGroupId;
                    }
                }
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.SP_LicenseProgram_Insert(EntityName, EntityNameEng, Code, StudyLevelId, AggregateGroupId, GSGUCode, idParam);
            int iVal = (int)idParam.Value;
            string query = "INSERT INTO SP_LicenseProgram (Id, Name, NameEng, Code, StudyLevelId, AggregateGroupId) VALUES (@Id, @Name, @NameEng, @Code, @StudyLevelId, @AggregateGroupId)";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", iVal);
            slParams.Add("@Name", EntityName);
            slParams.Add("@NameEng", EntityNameEng);
            slParams.Add("@Code", Code);
            slParams.Add("@StudyLevelId", StudyLevelId);
            slParams.Add("@AggregateGroupId", AggregateGroupId);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }

        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.SP_LicenseProgram_Update(EntityName, EntityNameEng, Code, StudyLevelId, AggregateGroupId, GSGUCode, id);
            string query = "UPDATE SP_LicenseProgram SET Name=@Name, NameEng=@NameEng, Code=@Code, StudyLevelId=@StudyLevelId, AggregateGroupId=@AggregateGroupId WHERE Id=@Id";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", id);
            slParams.Add("@Name", EntityName);
            slParams.Add("@NameEng", EntityNameEng);
            slParams.Add("@Code", Code);
            slParams.Add("@StudyLevelId", StudyLevelId);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }
    }
}
