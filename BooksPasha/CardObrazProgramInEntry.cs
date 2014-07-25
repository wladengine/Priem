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
    public partial class CardObrazProgramInEntry : BaseCard
    {
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
        
        public CardObrazProgramInEntry(Guid entryId, int licenseProgramId)
        {
            InitializeComponent();
            EntryId = entryId;
            LicenseProgramId = licenseProgramId;
            UpdateCombo();
            InitControls();
        }
        public CardObrazProgramInEntry(Guid Id)
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
                var z = context.ObrazProgramInEntry.Where(x => x.Id == GuidId).FirstOrDefault();
                if (z == null)
                {
                    WinFormsServ.Error("Не удалось получить значение ObrazProgramInEntry");
                    return;
                }
                EntryId = z.Entry.Id;
                LicenseProgramId = z.Entry.LicenseProgramId;
                UpdateCombo();
                ObrazProgramId = z.ObrazProgramId;
                KCP = z.KCP;
            }
            
            UpdateGridProfileInObrazProgramInEntry();
        }
        private void UpdateCombo()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var Ent = context.qObrazProgram.Where(x => x.LicenseProgramId == LicenseProgramId).Select(x => new { x.Id, x.Crypt, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Crypt + " " + x.Name)).ToList();
                ComboServ.FillCombo(cbObrazProgram, Ent, false, false);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveCard();
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            OpenCardProfileInObrazProgramInEntry(null);
        }

        private void OpenCardProfileInObrazProgramInEntry(string sId)
        {
            var crd = new CardProfileInObrazProgramInEntry(sId, tbLicenseProgram.Text, cbObrazProgram.Text);
            crd.OnSave += AddIntoBase;
            crd.Show();
        }

        private void UpdateGridProfileInObrazProgramInEntry()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.ProfileInObrazProgramInEntry.Where(x => x.ObrazProgramInEntryId == GuidId).Select(x => new { x.SP_Profile.Name, x.KCP }).ToArray();
                dgvProfileInObrazProgramInEntry.DataSource = Util.ConvertToDataTable(src);
            }
        }

        private void AddIntoBase(int ProfileId, int KCP)
        {
            if (!GuidId.HasValue)
                SaveCard();

            using (PriemEntities context = new PriemEntities())
            using (TransactionScope tran = new TransactionScope())
            {
                try
                {
                    Guid gId = Guid.NewGuid();
                    context.ProfileInObrazProgramInEntry.AddObject(new ProfileInObrazProgramInEntry() { ObrazProgramInEntryId = GuidId.Value, ProfileId = ProfileId, KCP = KCP, Id = gId });
                    context.SaveChanges();

                    string query = "INSERT INTO ProfileInObrazProgramInEntry (Id, ObrazProgramInEntryId, ProfileId, KCP) VALUES (@Id, @ObrazProgramInEntryId, @ProfileId, @KCP)";
                    SortedList<string, object> slParams = new SortedList<string, object>();
                    slParams.Add("@Id", gId);
                    slParams.Add("@ObrazProgramInEntryId", GuidId.Value);
                    slParams.Add("@ProfileId", ProfileId);
                    slParams.Add("@KCP", KCP);

                    MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
                    tran.Complete();
                }
                catch (Exception ex)
                {
                    WinFormsServ.Error(ex);
                }
            }
            UpdateGridProfileInObrazProgramInEntry();
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
                        context.ObrazProgramInEntry.AddObject(new ObrazProgramInEntry() { Id = GuidId.Value, ObrazProgramId = ObrazProgramId, KCP = KCP, EntryId = EntryId });

                        query = "INSERT INTO ObrazProgramInEntry(Id, ObrazProgramId, EntryId) VALUES (@Id, @ObrazProgramId, @EntryId)";
                    }
                    else
                    {
                        var Ent = context.ObrazProgramInEntry.Where(x => x.Id == GuidId).FirstOrDefault();
                        if (Ent == null)
                        {
                            WinFormsServ.Error("Не найдена запись в таблице ObrazProgramInEntry!");
                            return false;
                        }

                        Ent.ObrazProgramId = ObrazProgramId;
                        Ent.KCP = KCP;

                        query = "UPDATE ObrazProgramInEntry SET ObrazProgramId=@ObrazProgramId, EntryId=@EntryId WHERE Id=@Id";
                    }

                    context.SaveChanges();
                    
                    SortedList<string, object> slParams = new SortedList<string, object>();
                    slParams.Add("@Id", GuidId.Value);
                    slParams.Add("@ObrazProgramId", ObrazProgramId);
                    slParams.Add("@EntryId", EntryId);
                    MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);

                    tran.Complete();

                    if (ToUpdateList != null)
                        ToUpdateList();

                    return true;
                }
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
        
        private void dgvProfileInObrazProgramInEntry_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            string sId = dgvProfileInObrazProgramInEntry["Id", e.RowIndex].Value.ToString();

            OpenCardProfileInObrazProgramInEntry(sId);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvProfileInObrazProgramInEntry.SelectedCells.Count == 0)
                return;

            int rwInd = dgvProfileInObrazProgramInEntry.SelectedCells[0].RowIndex;
            Guid gId = (Guid)dgvProfileInObrazProgramInEntry["Id", rwInd].Value;

            using (PriemEntities context = new PriemEntities())
            {
                var ent = context.ProfileInObrazProgramInEntry.Where(x => x.Id == gId).FirstOrDefault();
                if (ent == null)
                    return;

                using (TransactionScope tran = new TransactionScope())
                {
                    try
                    {
                        context.ProfileInObrazProgramInEntry.DeleteObject(ent);
                        context.SaveChanges();

                        string query = "DELETE FROM ProfileInObrazProgramInEntry WHERE Id=@Id";
                        SortedList<string, object> slParams = new SortedList<string, object>();
                        slParams.Add("@Id", gId);

                        MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);

                        tran.Complete();
                    }
                    catch (Exception ex)
                    {
                        WinFormsServ.Error(ex);
                    }
                }
            }
            UpdateGridProfileInObrazProgramInEntry();
        }
    }
}
