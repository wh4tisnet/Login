using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
using Application = Microsoft.Office.Interop.Excel.Application;
using Button = System.Windows.Controls.Button;
using DataTable = System.Data.DataTable;
using TextBox = System.Windows.Controls.TextBox;
using Window = System.Windows.Window;

namespace Test1
{
    public static class GVars
    {
        public static string filePath = @"C:\Users\abel.alvarez\Desktop\database.xlsx";
    }

    public partial class MainWindow
    {
        DataTable table = new DataTable();

        public MainWindow()
        {
            InitializeComponent();
            CreateTable();
        }

        private void OpenPart2_Click(object sender, RoutedEventArgs e)
        {
            part2 newWindow = new part2();
            newWindow.Show();
        }

        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            string username = UsernameTextBox.Text;
            string password = PasswordBox.Password;

            DataRow[] foundRows = table.Select("Name = '" + username + "' AND Password = '" + password + "'");

            if (foundRows.Length > 0)
            {
                MessageBox.Show("Successful login for: " + username, "Nice", MessageBoxButton.OK, MessageBoxImage.Information);
                DataRow userRow = foundRows[0];
                accountmanagemnt accountwindow = new accountmanagemnt();
                accountwindow.information = userRow;
                accountwindow.ShowDialog();
            }
            else
            {
                MessageBox.Show("Incorrect username or password", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void ShowPassword_Checked(object sender, RoutedEventArgs e)
        {
            VisiblePasswordBox.Text = PasswordBox.Password;
            VisiblePasswordBox.Visibility = Visibility.Visible;
        }

        private void ShowPassword_Unchecked(object sender, RoutedEventArgs e)
        {
            PasswordBox.Password = VisiblePasswordBox.Text;
            VisiblePasswordBox.Visibility = Visibility.Collapsed;
        }

        private void CreateTable()
        {
            table = ReadExcelFile(GVars.filePath);
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {

            string currentUsername = UsernameTextBox.Text;

            Window resetPasswordWindow = new Window
            {
                Title = "Reset Password",
                Width = 300,
                Height = 250,
            };

            StackPanel stackPanel = new StackPanel { Margin = new Thickness(10) };

            TextBox usernameBox = new TextBox { Width = 200, IsReadOnly = true, Text = currentUsername };
            PasswordBox newPasswordBox = new PasswordBox { Width = 200 };
            PasswordBox repeatPasswordBox = new PasswordBox { Width = 200 };

            Button confirmButton = new Button { Content = "Confirm", Width = 80 };

            stackPanel.Children.Add(new TextBlock { Text = "Username:" });
            stackPanel.Children.Add(usernameBox);
            stackPanel.Children.Add(new TextBlock { Text = "New Password:" });
            stackPanel.Children.Add(newPasswordBox);
            stackPanel.Children.Add(new TextBlock { Text = "Repeat New Password:" });
            stackPanel.Children.Add(repeatPasswordBox);
            stackPanel.Children.Add(confirmButton);

            resetPasswordWindow.Content = stackPanel;

            confirmButton.Click += (s, args) =>
            {
                string username = usernameBox.Text;
                string newPassword = newPasswordBox.Password;
                string repeatPassword = repeatPasswordBox.Password;

                if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(newPassword) || string.IsNullOrEmpty(repeatPassword))
                {
                    MessageBox.Show("Enter a correct user.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (newPassword != repeatPassword)
                {
                    MessageBox.Show("Passwords do not match.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    DataRow[] foundRows = table.Select($"Name = '{username}'");
                    if (foundRows.Length > 0)
                    {
                        foundRows[0]["Password"] = newPassword;
                        MessageBox.Show($"Password for user '{username}' successfully reset.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        SaveTableToExcel(GVars.filePath, table);
                        resetPasswordWindow.Close();
                    }
                    else
                    {
                        MessageBox.Show($"User '{username}' not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            };

            resetPasswordWindow.ShowDialog();
        }

        private static DataTable ReadExcelFile(string filePath)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1];
            Range range = worksheet.UsedRange;

            DataTable dataTable = new DataTable();

            for (int col = 1; col <= range.Columns.Count; col++)
            {
                string columnName = (range.Cells[1, col] as Range).Value2?.ToString();
                if (string.IsNullOrEmpty(columnName))
                {
                    columnName = $"Column{col}";
                }
                dataTable.Columns.Add(columnName);
            }

            int row = 2;
            while ((range.Cells[row, 1] as Range).Value2 != null)
            {
                DataRow dataRow = dataTable.NewRow();

                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dataRow[col - 1] = (range.Cells[row, col] as Range).Value2;
                }
                
                dataTable.Rows.Add(dataRow);
                row++;
            }

            workbook.Close(true);
            excelApp.Quit();

            return dataTable;
        }

        private static void SaveTableToExcel(string filePath, DataTable dataTable)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1];

            worksheet.Cells.Clear();

            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                worksheet.Cells[1, col + 1] = dataTable.Columns[col].ColumnName;
            }

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1] = dataTable.Rows[row][col].ToString();
                }
            }

            workbook.Save();
            workbook.Close(true);
            excelApp.Quit();


        }
    }
}
