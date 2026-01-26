using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace IskurTakipSistemi.DataBase.Repos
{
    internal class PuantajRepo
    {
        private readonly DbRepository _db = new DbRepository();
        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM puantaj ");
        }
    }
}
