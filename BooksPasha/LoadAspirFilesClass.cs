using PriemLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Priem
{
    public static class LoadAspirFilesClass
    {
        public static void LoadFiles()
        {
            FolderBrowserDialog fld = new FolderBrowserDialog();
            var dr = fld.ShowDialog();

            if (dr == DialogResult.OK)
            {
                string dir = fld.SelectedPath;
                
                List<string> lstOP = new List<string>();
                lstOP.Add("Математическая кибернетика");
                lstOP.Add("Системный анализ, информатика и управление");
                lstOP.Add("Информатика");
                lstOP.Add("Механика");
                lstOP.Add("Прикладная математика и процессы управления");
                foreach (string OP in lstOP)
                {
                    string query = @"SELECT qq.Id, F.FileExtention, F.FileData
FROM OnlineAbitFiles.dbo.qAbitFiles_OnlyEssay qq
INNER JOIN extAbitFiles_All F ON qq.Id = F.Id
WHERE qq.PersonId IN
(
	SELECT DISTINCT APP.PersonId
	FROM Application2015 APP
	INNER JOIN Entry E ON APP.EntryId = E.Id
	WHERE E.StudyLevelGroupId = 4
	AND E.ObrazProgramName = @OP
)
UNION ALL
SELECT qq.Id, F.FileExtention, F.FileData
FROM OnlinePriem2015.dbo.qAbitFiles_OnlyEssay qq
INNER JOIN extAbitFiles_All F ON qq.Id = F.Id
WHERE qq.PersonId IN
(
	SELECT DISTINCT APP.PersonId
	FROM Application2015 APP
	INNER JOIN Entry E ON APP.EntryId = E.Id
	WHERE E.StudyLevelGroupId = 4
	AND E.ObrazProgramName = @OP
)";
                    DataTable tbl = MainClass.BdcOnlineReadWrite.GetDataSet(query, new SortedList<string, object>() { { "@OP", OP } }).Tables[0];
                    foreach (DataRow rw in tbl.Rows)
                    {
                        string fileName = dir + "\\(" + OP + ")" + rw["Id"].ToString() + rw["FileExtention"].ToString();
                        byte[] fileData = rw.Field<byte[]>("FileData");

                        using (var fs = System.IO.File.Create(fileName))
                        {
                            fs.Write(fileData, 0, fileData.Length);
                            fs.Close();
                        }
                    }
                }
            }
        }
    }
}
