using System;
using Microsoft.Data.Sqlite;
using System.Data;

namespace denemelikimid.DataBase
{
    public class NOTDbRepository
    {
        public DataTable GetAll(string tableName)
        {
            try
            {
                var datatable = new DataTable();

                using (var con = NotDbConnection.GetConnection())
                {
                    if (con.State != ConnectionState.Open)
                        con.Open();

                    using (var cmd = new SqliteCommand($"SELECT * FROM {tableName}", con))
                    using (var reader = cmd.ExecuteReader())
                    {
                        datatable.Load(reader);
                    }
                }
                return datatable;
            }
            catch (Exception ex)
            {
                throw new Exception("Başarısız veri çekme işlemi, hata: " + tableName, ex);
            }
        }

        public DataTable GetByQuery(string sql, params SqliteParameter[] parameters)
        {
            try
            {
                var datatable = new DataTable();

                using (var con = NotDbConnection.GetConnection())
                {
                    using (var cmd = new SqliteCommand(sql, con))
                    {
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        if (con.State != ConnectionState.Open)
                            con.Open();

                        using (var reader = cmd.ExecuteReader())
                        {
                            datatable.Load(reader);
                        }
                    }
                }

                return datatable;
            }
            catch (Exception ex)
            {
                throw new Exception("Başarısız sorgu işlemi, hata: " + sql, ex);
            }
        }

        public void Execute(string sql, params SqliteParameter[] parameters)
        {
            try
            {
                using (var con = NotDbConnection.GetConnection())
                {
                    using (var cmd = new SqliteCommand(sql, con))
                    {
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        if (con.State != ConnectionState.Open)
                            con.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Başarısız veri tabanı komutu, hata: " + sql, ex);
            }
        }
    }
}

