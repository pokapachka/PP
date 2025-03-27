using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Prog.Connection;
using Prog.Models;

namespace Prog.Pages
{
    /// <summary>
    /// Логика взаимодействия для Auth.xaml
    /// </summary>
    public partial class Auth : Page
    {
        public Auth()
        {
            InitializeComponent();
        }

        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string username = Login.Text;
            string password = Password.Password;
            string dbName = (NameDB.SelectedItem as ComboBoxItem)?.Content.ToString();
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password) || string.IsNullOrEmpty(dbName))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }
            try
            {
                Model.Username = username;
                Model.Password = password;
                Model.DbName = dbName;
                Model.CurrentConnection = new ConnectionDB(dbName, username, password);
                if (Model.CurrentConnection != null && Model.CurrentConnection.GetConnection()?.State == ConnectionState.Open)
                {
                    MainWindow.mainWindow.frame.Navigate(new Main());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка авторизации: {ex.Message}");
                Model.CurrentConnection = null;
            }
        }
        
    }
}
