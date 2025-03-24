using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;

namespace Prog.Connection
{
    public class ConnectionDB
    {
        private OleDbConnection connection;
        private const string ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=MyOracleDB;User Id=username;Password=password;";
        public ConnectionDB()
        {
            try
            {
                connection = new OleDbConnection(ConnectionString);
                connection.Open();
                MessageBox.Show("Подключение к Oracle успешно установлено.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при подключении к Oracle: {ex.Message}");
            }
        }

        public OleDbConnection GetConnection()
        {
            return connection;
        }

        public void CloseConnection()
        {
            if (connection != null && connection.State == ConnectionState.Open)
            {
                connection.Close();
                MessageBox.Show("Соединение с Oracle закрыто.");
            }
        }
    }
}