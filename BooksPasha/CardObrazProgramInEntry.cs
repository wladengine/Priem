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
using System.Transactions;

namespace Priem
{
    public partial class CardInnerEntryInEntry : BaseCard
    {
        #region Fields
        public event UpdateListHandler ToUpdateList;
        private Guid? GuidId;
        private Guid EntryId;
        private int _licenseProgramId;
        private int LicenseProgramId 
        { 
            get { return _licenseProgramId; }
            set
            {
                using (PriemEntities context = new PriemEntities())
                {
                    _licenseProgramId = value;
                    tbLicenseProgram.Text = context.SP_LicenseProgram.Where(x => x.Id == _licenseProgramId)
                        .Select(x => new { x.Code, x.Name }).ToList().Select(x => x.Code + " " + x.Name).FirstOrDefault();
                }
            }
        }
        private int KCP
        {
            get
            {
                int iRet = 0;
                int.TryParse(tbKCP.Text, out iRet);
                return iRet;
            }
            set { tbKCP.Text = value.ToString(); }
        }
        private int ObrazProgramId
        {
            get { return ComboServ.GetComboIdInt(cbObrazProgram).Value; }
            set { ComboServ.SetComboId(cbObrazProgram, value); }
        }
        private int ProfileId
        {
            get { return ComboServ.GetComboIdInt(cbProfile).Value; }
            set { ComboServ.SetComboId(cbProfile, value); }
        }
        #endregion

        public CardInnerEntryInEntry(Guid entryId, int licenseProgramId)
        {
            InitializeComponent();
            EntryId = entryId;
            LicenseProgramId = licenseProgramId;
            UpdateCombo();
            InitControls();
        }
        public CardInnerEntryInEntry(Guid Id)
        {
            InitializeComponent();
            GuidId = Id;
            LoadValues();
            InitControls();
        }

        protected override void ExtraInit()
        {
            this.MdiParent = MainClass.mainform;
        }
        protected override void FillCard()
        {
            LoadValues();
        }
        private void LoadValues()
        {
            if (!GuidId.HasValue)
                return;
            using (PriemEntities context = new PriemEntities())
            {
                var z = context.InnerEntryInEntry.Where(x => x.Id == GuidId).FirstOrDefault();
                if (z == null)
                {
                    WinFormsServ.Error("Не удалось получить значение InnerEntryInEntry");
                    return;
                }

                EntryId = z.Entry.Id;
                LicenseProgramId = z.Entry.LicenseProgramId;
                UpdateCombo();
                ObrazProgramId = z.ObrazProgramId;
                ProfileId = z.ProfileId;
                KCP = z.KCP;
            }
        }
        private void UpdateCombo()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var Ent = context.qObrazProgram.Where(x => x.LicenseProgramId == LicenseProgramId).Select(x => new { x.Id, x.Crypt, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Crypt + " " + x.Name)).ToList();
                ComboServ.FillCombo(cbObrazProgram, Ent, false, false);

                var Prof = context.SP_Profile.Select(x => new { x.Id, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbProfile, Prof, false, false);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveCard();
        }
        
        protected override bool SaveRecord()
        {
            return SaveCard();
        }
        private bool SaveCard()
        {
            try
            {
                using (TransactionScope tran = new TransactionScope())
                using (PriemEntities context = new PriemEntities())
                {
                    string query = "";
                    if (!GuidId.HasValue)
                    {
                        GuidId = Guid.NewGuid();
                        context.InnerEntryInEntry.AddObject(new InnerEntryInEntry() { Id = GuidId.Value, ObrazProgramId = ObrazProgramId, ProfileId = ProfileId, KCP = KCP, EntryId = EntryId });

                        query = "INSERT INTO InnerEntryInEntry (Id, ObrazProgramId, ProfileId, EntryId) VALUES (@Id, @ObrazProgramId, @ProfileId, @EntryId)";
                    }
                    else
                    {
                        var Ent = context.InnerEntryInEntry.Where(x => x.Id == GuidId).FirstOrDefault();
                        if (Ent == null)
                        {
                            WinFormsServ.Error("Не найдена запись в таблице InnerEntryInEntry!");
                            return false;
                        }

                        Ent.ObrazProgramId = ObrazProgramId;
                        Ent.ProfileId = ProfileId;
                        Ent.KCP = KCP;

                        query = "UPDATE InnerEntryInEntry SET ObrazProgramId=@ObrazProgramId, ProfileId=@ProfileId, EntryId=@EntryId WHERE Id=@Id";
                    }

                    context.SaveChanges();
                    
                    SortedList<string, object> slParams = new SortedList<string, object>();
                    slParams.Add("@Id", GuidId.Value);
                    slParams.Add("@ObrazProgramId", ObrazProgramId);
                    slParams.Add("@EntryId", EntryId);
                    slParams.Add("@ProfileId", ProfileId);
                    MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);

                    tran.Complete();
                }

                if (ToUpdateList != null)
                    ToUpdateList();

                return true;
            }
            catch (Exception ex)
            {
                WinFormsServ.Error(ex);
                return false;
            }
        }

        protected override void CloseCardAfterSave()
        {
            this.Close();
        }
    }
}
