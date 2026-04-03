using Microsoft.Data.Sqlite;
using System;
using System.IO;

namespace denemelikimid.DataBase
{
    public static class DbConnection
    {
        private static readonly string _connectionString =
            new SqliteConnectionStringBuilder
            {
                DataSource = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "iskur.db")
            }.ToString();

        public static SqliteConnection GetConnection()
        {
            return new SqliteConnection(_connectionString);
        }
    }
}
