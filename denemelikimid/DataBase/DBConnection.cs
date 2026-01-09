using System;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace denemelikimid.DataBase
{
    public static class DbConnection
    {
        private static readonly string _connectionString =
            "Server=localhost;Database=sakila;Uid=yeniAdmin;Pwd=1234;";

        public static MySqlConnection GetConnection()
        {
            MySqlConnection connection = new MySqlConnection(_connectionString);

            try
            {
                connection.Open();
                return connection;
            }
            catch (MySqlException ex)
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

