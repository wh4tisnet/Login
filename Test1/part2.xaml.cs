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
using System.Windows.Media.Animation;
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
    public partial class part2
    {
        DataTable table = new DataTable();

        public part2()
        {
            InitializeComponent();
            CreateTable();
        }

        private void CreateTable()
        {
            table = ReadExcelFile(GVars.filePath);
            bs.ItemsSource = table.DefaultView;
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

        private void Createuser_Click(object sender, RoutedEventArgs e)
        {
            DataRow userNew = table.NewRow();
            accountmanagemnt accountwindow = new accountmanagemnt();
            accountwindow.information = userNew;
            accountwindow.ShowDialog();
            userNew = accountwindow.information;
            table.Rows.Add(userNew);
        }


        private void DeleteUser_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = (DataRowView)bs.SelectedItem;
            if (selectedItem != null)
            {
                int idToDelete = int.Parse(selectedItem.Row["Id"].ToString());
                DeleteUser_Click(idToDelete);
            }
            else
            {
                MessageBox.Show("Please select a user from the list.");
            }
        }

        private void UpdateTable_Click(object sender, RoutedEventArgs e)
        {
            table = ReadExcelFile(GVars.filePath);

            bs.ItemsSource = table.DefaultView;

            MessageBox.Show("Table updated successfully.", "Success", MessageBoxButton.OK);
        }

        private void DeleteUser_Click(int userId)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(GVars.filePath);
            Worksheet worksheet = workbook.Sheets[1];
            Range range = worksheet.UsedRange;

            bool userFound = false;
            int row = 2;

            while ((range.Cells[row, 1] as Range).Value2 != null)
            {
                var cellValue = Convert.ToString((range.Cells[row, 1] as Range).Value2);
                if (int.TryParse(cellValue, out int currentUserId) && currentUserId == userId)
                {
                    worksheet.Rows[row].Delete();
                    userFound = true;
                    break;
                }
                row++;
            }

            if (userFound)
            {
                int currentRow = 2;
                while ((range.Cells[currentRow, 1] as Range).Value2 != null)
                {
                    worksheet.Cells[currentRow, 1].Value2 = currentRow - 1;
                    currentRow++;
                }

                workbook.Save();
                MessageBox.Show("User deleted and IDs reorganized successfully.", "Success", MessageBoxButton.OK);
            }
            else
            {
                MessageBox.Show("No users with the specified ID were found.", "User Not Found", MessageBoxButton.OK);
            }

            workbook.Close(true);
            excelApp.Quit();

        }

        private void Editsuser_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = (DataRowView)bs.SelectedItem;

            if (selectedItem != null)
            {
                DataRow selectedUser = selectedItem.Row;
                accountmanagemnt accountWindow = new accountmanagemnt();
                accountWindow.information = selectedUser;
                accountWindow.ShowDialog();
                selectedUser = accountWindow.information;

                if (selectedUser != null)
                {
                    bs.ItemsSource = table.DefaultView;
                }
            }
            else
            {
                MessageBox.Show("Please select a user from the list to edit.");
            }
        }
    }
}
