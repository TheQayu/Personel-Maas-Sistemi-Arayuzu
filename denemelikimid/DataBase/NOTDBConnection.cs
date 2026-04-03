using System;
using Microsoft.Data.Sqlite;
using System.Windows.Forms;
using System.Configuration;
using System.IO;

namespace denemelikimid.DataBase
{
    public static class NotDbConnection
    {
        private static readonly string _connectionString =
            ConfigurationManager.ConnectionStrings["IskurDb"]?.ConnectionString
            ?? new SqliteConnectionStringBuilder
            {
                DataSource = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "iskur.db")
            }.ToString();

        public static SqliteConnection GetConnection()
        {
            SqliteConnection connection = new SqliteConnection(_connectionString);

            try
            {
                connection.Open();
                return connection;
            }
            catch (SqliteException ex)
            {
                MessageBox.Show("Veritabanı bağlantı hatası: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Genel bağlantı hatası: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }
}

