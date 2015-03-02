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
    public partial class CardObrazProgram : BookCardInt
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
        private string Number
        {
            get { return tbNumber.Text.Trim(); }
            set { tbNumber.Text = value; }
        }
        private int FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty).Value; }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        private int LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram).Value; }
            set { ComboServ.SetComboId(cbLicenseProgram, value); }
        }
        private int? StudyLevelId;
        public CardObrazProgram(string id, int? iFacultyId, int? iStudyLevelId)
            : base(id)
        {
            InitializeComponent();
            StudyLevelId = iStudyLevelId;
            InitControls();
            if (iFacultyId.HasValue)
                FacultyId = iFacultyId.Value;
        }

        protected override void ExtraInit()
        {
            this.MdiParent = MainClass.mainform;
            _title = "Образовательная программа";
            _tableName = "ed.SP_ObrazProgram";
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.SP_Faculty.Select(x => new { x.Id, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbFaculty, src, false, false);

                src = context.SP_LicenseProgram.Where(x => StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true)
                    .Select(x => new { x.Id, x.Code, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), "(" + x.Code + ") " + x.Name)).OrderBy(x => x.Value).ToList();
                ComboServ.FillCombo(cbLicenseProgram, src, false, false);
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
                    var OP = context.SP_ObrazProgram.Where(x => x.Id == IntId).FirstOrDefault();
                    if (OP == null)
                        WinFormsServ.Error("Не найдена запись в базе!");
                    else
                    {
                        EntityName = OP.Name;
                        EntityNameEng = OP.NameEng;
                        Number = OP.Number;
                        FacultyId = OP.FacultyId;
                        LicenseProgramId = OP.LicenseProgramId;
                    }
                }
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.SP_ObrazProgram_Insert(EntityName, EntityNameEng, Number, FacultyId, LicenseProgramId, idParam);

            int iVal = (int)idParam.Value;
            string query = "INSERT INTO SP_ObrazProgram (Id, Name, NameEng, Number, FacultyId, LicenseProgramId, ProgramModeId) VALUES (@Id, @Name, @NameEng, @Number, @FacultyId, @LicenseProgramId, @ProgramModeId)";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", iVal);
            slParams.Add("@Name", EntityName);
            slParams.Add("@NameEng", EntityNameEng);
            slParams.Add("@Number", Number);
            slParams.Add("@FacultyId", FacultyId);
            slParams.Add("@LicenseProgramId", LicenseProgramId);
            slParams.Add("@ProgramModeId", 1);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }

        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.SP_ObrazProgram_Update(EntityName, EntityNameEng, Number, FacultyId, LicenseProgramId, id);

            string query = "UPDATE SP_ObrazProgram SET Name=@Name, NameEng=@NameEng, Number=@Number, FacultyId=@FacultyId, LicenseProgramId=@LicenseProgramId WHERE Id=@Id";
            SortedList<string, object> slParams = new SortedList<string, object>();
            slParams.Add("@Id", id);
            slParams.Add("@Name", EntityName);
            slParams.Add("@NameEng", EntityNameEng);
            slParams.Add("@Number", Number);
            slParams.Add("@FacultyId", FacultyId);
            slParams.Add("@LicenseProgramId", LicenseProgramId);
            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
        }
    }
}