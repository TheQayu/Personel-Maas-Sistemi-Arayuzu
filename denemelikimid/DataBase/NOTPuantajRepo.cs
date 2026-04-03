using System;
using System.Data;
using Microsoft.Data.Sqlite;

namespace denemelikimid.DataBase
{
    public class NOTPuantajRepo
    {
        private readonly NOTDbRepository _db = new NOTDbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM puantaj ");
        }

        public DataTable GetByKatilimci(double tc_no)
        {
            return _db.GetByQuery(
                @"SELECT * FROM puantaj WHERE p_tc_no = @tc",
                new SqliteParameter("@tc", tc_no));
        }

        public void Ekle(double tc_no, int gun, string durum)
        {
            _db.Execute(
                @"INSERT OR REPLACE INTO puantaj (p_tc_no, p_gun, p_durum) VALUES (@tc, @gun, @durum)",
                new SqliteParameter("@tc", tc_no),
                new SqliteParameter("@gun", gun),
                new SqliteParameter("@durum", durum));
        }
    }
}







