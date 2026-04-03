using Microsoft.Data.Sqlite;
using System;
using System.Data;

namespace denemelikimid.DataBase
{
    public class DbRepository
    {
        // Parametreli select sorguları için DataTable döndüren method
        public DataTable GetByQuery(string sql, params SqliteParameter[] parameters)
        {
            try
            {
                var dt = new DataTable();

                using (var con = DbConnection.GetConnection())
                {
                    using (var cmd = new SqliteCommand(sql, con))
                    {
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        if (con.State != ConnectionState.Open)
                            con.Open();

                        using (var reader = cmd.ExecuteReader())
                        {
                            dt.Load(reader);
                        }
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                
                throw;
            }
        }

        // Parametreli sorgular için (Insert, Update, Delete)
        public void Execute(string sql, params SqliteParameter[] parameters)
        {
            try
            {
                using (var con = DbConnection.GetConnection())
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
            catch (Exception)
            {
                throw;
            }
        }
    }
}