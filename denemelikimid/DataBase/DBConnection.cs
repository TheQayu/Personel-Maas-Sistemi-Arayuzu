using MySql.Data.MySqlClient;

namespace IskurTakipSistemi.DataBase
{
    public static class DbConnection
    {
        private static readonly string _connectionString =
            "Server=localhost;Database=iskur;Uid=root;Pwd=ms123456.";

        public static MySqlConnection GetConnection()
        {
            return new MySqlConnection(_connectionString);
        }
    }
}
