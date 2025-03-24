﻿using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;

namespace Prog.Connection
{
    /// <summary>
    /// Подключение к БД
    /// </summary>
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
                MessageBox.Show("Подключение успешно установлено.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при подключении: {ex.Message}");
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
                MessageBox.Show("Соединение закрыто.");
            }
        }
        /// <summary>
        /// Создание временной таблицы и занесение туда верных строк
        /// </summary>
        public void ProcessGreenRecords()
        {
            try
            {
                // 1. Создаем временную таблицу
                var cmdCreateTemp = new OleDbCommand(@"CREATE GLOBAL TEMPORARY TABLE temp_green_records (id NUMBER, op_kzr VARCHAR2(100), op_dse VARCHAR2(100), cnt NUMBER) ON COMMIT PRESERVE ROWS", connection);
                cmdCreateTemp.ExecuteNonQuery();
                // 2. Копируем "зелёные" записи
                var cmdInsertGreen = new OleDbCommand(@"INSERT INTO temp_green_records (id, op_kzr, op_dse, cnt) SELECT id, op_kzr, op_dse, cnt FROM main_table WHERE flag = 0", connection);
                int insertedRows = cmdInsertGreen.ExecuteNonQuery();
                // 3. Сравниваем и добавляем новые записи (через MERGE)
                var cmdMerge = new OleDbCommand(@"MERGE INTO permanent_table p USING temp_green_records t ON (p.op_kzr = t.op_kzr AND p.op_dse = t.op_dse) WHEN NOT MATCHED THEN INSERT (op_kzr, op_dse, cnt) VALUES (t.op_kzr, t.op_dse, t.cnt)", connection);
                int mergedRows = cmdMerge.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
        }
    }
}