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
    public static class FilestreamTestClass
    {
        public static void TestHashFunc()
        {
            DBPriem _bdcOnlineReadWrite = new DBPriem();

            Int64 iBytes = 0;
            _bdcOnlineReadWrite.OpenDatabase(DBConstants.CS_PriemONLINE_Files);

            DateTime dtStart = DateTime.Now;

            string query = "SELECT TOP 1 Id, FileData FROM FileStorage";
            DataTable tbl = _bdcOnlineReadWrite.GetDataSet(query).Tables[0];

            TimeSpan ts1 = DateTime.Now - dtStart;
            dtStart = DateTime.Now;

            foreach (DataRow rw in tbl.Rows)
            {
                Guid FileId = rw.Field<Guid>("Id");
                byte[] FileData = rw.Field<byte[]>("FileData");
                iBytes += FileData.Length;
                string sHash = SHA1Byte(FileData);
            }
            TimeSpan ts2 = DateTime.Now - dtStart;
            MessageBox.Show(tbl.Rows.Count + " файлов, " + iBytes + " байт, на запрос " + ts1.ToString() + ", на хэширование " + ts2.ToString());
        }


        /// <summary>
        /// Возвращает SHA1-строку от byte[] источника
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static string SHA1Byte(byte[] source)
        {
            byte[] md5 = System.Security.Cryptography.SHA1.Create().ComputeHash(source);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in md5)
            {
                sb.Append(b.ToString("x2"));
            }

            return sb.ToString();
        }
    }
}
