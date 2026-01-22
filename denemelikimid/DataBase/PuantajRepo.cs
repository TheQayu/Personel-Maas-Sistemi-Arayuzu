using System;
using System.Data;
using MySql.Data.MySqlClient;

namespace denemelikimid.DataBase
{
    public class PuantajRepo
    {
        private readonly DbRepository _db = new DbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM puantaj ");
        }

        public DataTable GetByKatilimci(double tc_no)
        {
            return _db.GetByQuery(
                @"SELECT * FROM puantaj WHERE p_tc_no = @tc",
                new MySqlParameter("@tc", tc_no));
        }

        public void Ekle(double tc_no, int gun, string durum)
        {
            _db.Execute(
                @"INSERT INTO puantaj (p_tc_no, p_gun, p_durum) VALUES (@tc, @gun, @durum) " +
                "ON DUPLICATE KEY UPDATE p_durum = @durum",
                new MySqlParameter("@tc", tc_no),
                new MySqlParameter("@gun", gun),
                new MySqlParameter("@durum", durum));
        }
    }
}







