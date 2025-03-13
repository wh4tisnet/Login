using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Data.Common;

namespace Test1
{
    public partial class accountmanagemnt
    {
        private readonly string filePath;

        public DataRow information { get; set; }
        public static object userData { get; private set; }

        public accountmanagemnt()
        {
            InitializeComponent();
            DataTable excelData = ReadExcelFile(GVars.filePath, GetUserData());
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            IdTextBox.Text = information["Id"].ToString();
            NameTextBox.Text = information["Name"].ToString();
            PasswordBox.Password = information["Password"].ToString();
            EmailTextBox.Text = information["Email"].ToString();
            LiveTextBox.Text = information["Live"].ToString();
            VideogameTextBox.Text = information["Videogames"].ToString();
            PhoneTextBox.Text = information["Phone"].ToString();
            AddressTextBox.Text = information["Address"].ToString();
            CountryTextBox.Text = information["Country"].ToString();
            AgeTextBox.Text = information["Age"].ToString();
            FavoriteMovieTextBox.Text = information["Favorite Movie"].ToString();
            UsersTextBox.Text = information["Users"].ToString();
        }

        private void Saveas_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateFields())
            {
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(GVars.filePath);
                Worksheet worksheet = workbook.Sheets[1];
                Range range = worksheet.UsedRange;

                InsertNewRow(worksheet);
                ReorganizeIds(worksheet);

                workbook.Save();
                workbook.Close();
                excelApp.Quit();
            }
            else
            {
                MessageBox.Show("Please fill out all fields correctly.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ValidateFields()
        {
            if (string.IsNullOrEmpty(NameTextBox.Text) || string.IsNullOrEmpty(PasswordBox.Password))
            {
                MessageBox.Show("Name and Password cannot be empty.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (!IsValidEmail(EmailTextBox.Text))
            {
                MessageBox.Show("Please enter a valid email address.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (!IsValidPhoneNumber(PhoneTextBox.Text))
            {
                MessageBox.Show("Phone number must contain only numbers.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (!IsValidAge(AgeTextBox.Text))
            {
                MessageBox.Show("Age must be a valid positive number.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private bool IsValidEmail(string email)
        {
            return email.Contains("@") && email.Contains(".");
        }

        private bool IsValidPhoneNumber(string phone)
        {
            return !string.IsNullOrEmpty(phone) && phone.All(char.IsDigit);
        }

        private bool IsValidAge(string age)
        {
            return int.TryParse(age, out int result) && result > 0;
        }

        private void InsertNewRow(Worksheet worksheet)
        {
            Range range = worksheet.UsedRange;
            int emptyRowIndex = range.Rows.Count + 1;

            int newId = GenerateNewId(worksheet);
            worksheet.Cells[emptyRowIndex, 1].Value = newId;
            worksheet.Cells[emptyRowIndex, 2].Value = NameTextBox.Text;
            worksheet.Cells[emptyRowIndex, 3].Value = PasswordBox.Password;
            worksheet.Cells[emptyRowIndex, 4].Value = EmailTextBox.Text;
            worksheet.Cells[emptyRowIndex, 5].Value = LiveTextBox.Text;
            worksheet.Cells[emptyRowIndex, 6].Value = VideogameTextBox.Text;
            worksheet.Cells[emptyRowIndex, 7].Value = PhoneTextBox.Text;
            worksheet.Cells[emptyRowIndex, 8].Value = AddressTextBox.Text;
            worksheet.Cells[emptyRowIndex, 9].Value = CountryTextBox.Text;
            worksheet.Cells[emptyRowIndex, 10].Value = AgeTextBox.Text;
            worksheet.Cells[emptyRowIndex, 11].Value = FavoriteMovieTextBox.Text;
            worksheet.Cells[emptyRowIndex, 11].Value = UsersTextBox.Text;
        }

        private void UpdateRow(Worksheet worksheet, int row)
        {
            worksheet.Cells[row, 2].Value = NameTextBox.Text;
            worksheet.Cells[row, 3].Value = PasswordBox.Password;
            worksheet.Cells[row, 4].Value = EmailTextBox.Text;
            worksheet.Cells[row, 5].Value = LiveTextBox.Text;
            worksheet.Cells[row, 6].Value = VideogameTextBox.Text;
            worksheet.Cells[row, 7].Value = PhoneTextBox.Text;
            worksheet.Cells[row, 8].Value = AddressTextBox.Text;
            worksheet.Cells[row, 9].Value = CountryTextBox.Text;
            worksheet.Cells[row, 10].Value = AgeTextBox.Text;
            worksheet.Cells[row, 11].Value = FavoriteMovieTextBox.Text;
            worksheet.Cells[row, 11].Value = UsersTextBox.Text;
        }

        private int GenerateNewId(Worksheet worksheet)
        {
            Range range = worksheet.UsedRange;
            int lastId = 0;
            int row = 1;

            while ((range.Cells[row, 1] as Range).Value2 != null)
            {
                int currentId;
                if (int.TryParse(worksheet.Cells[row, 1].Value?.ToString(), out currentId))
                {
                    if (currentId > lastId)
                    {
                        lastId = currentId;
                    }
                }
                row++;
            }

            return lastId + 1;
        }

        private void ReorganizeIds(Worksheet worksheet)
        {
            Range range = worksheet.UsedRange;
            int newId = 1;

            for (int row = 2; row <= range.Rows.Count; row++)
            {
                bool isRowNotEmpty = false;

                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    if ((range.Cells[row, col] as Range).Value2 != null)
                    {
                        isRowNotEmpty = true;
                        break;
                    }
                }

                if (isRowNotEmpty)
                {
                    worksheet.Cells[row, 1].Value = newId;
                    newId++;
                }
                else
                {
                    worksheet.Rows[row].Delete();
                    row--;
                }
            }
        }

        private void DeleteUser_Click(object sender, RoutedEventArgs e)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(GVars.filePath);
            Worksheet worksheet = workbook.Sheets[1];
            Range range = worksheet.UsedRange;

            if (information["Id"] != null)
            {
                string idToFind = information["Id"].ToString();
                bool isFound = false;

                int row = 1;
                while ((range.Cells[row, 1] as Range).Value2 != null)
                {
                    if (worksheet.Cells[row, 1].Value?.ToString() == idToFind)
                    {
                        isFound = true;
                        DeleteRow(worksheet, row);
                        break;
                    }
                    row++;
                }

                if (!isFound)
                {
                    MessageBox.Show("User ID not found.");
                }
            }
            else
            {
                MessageBox.Show("User ID is null.");
            }

            ReorganizeIds(worksheet);

            workbook.Save();
            workbook.Close();
            excelApp.Quit();
        }

        private void DeleteRow(Worksheet worksheet, int row)
        {
            worksheet.Rows[row].Delete();
        }

        private void InsertInEmptyRow(Worksheet worksheet)
        {
            int emptyRowIndex = 1;

            while (worksheet.Cells[emptyRowIndex, 1].Value != null)
                emptyRowIndex++;

            worksheet.Cells[emptyRowIndex, 1] = NameTextBox.Text;
            worksheet.Cells[emptyRowIndex, 2] = PasswordBox.Password;
            worksheet.Cells[emptyRowIndex, 3] = EmailTextBox.Text;
            worksheet.Cells[emptyRowIndex, 4] = LiveTextBox.Text;
            worksheet.Cells[emptyRowIndex, 5] = VideogameTextBox.Text;
            worksheet.Cells[emptyRowIndex, 6] = PhoneTextBox.Text;
            worksheet.Cells[emptyRowIndex, 7] = AddressTextBox.Text;
            worksheet.Cells[emptyRowIndex, 8] = CountryTextBox.Text;
            worksheet.Cells[emptyRowIndex, 9] = AgeTextBox.Text;
            worksheet.Cells[emptyRowIndex, 10] = FavoriteMovieTextBox.Text;
            worksheet.Cells[emptyRowIndex, 11] = UsersTextBox;
        }

        private static object GetUserData()
        {
            return userData;
        }

        private static DataTable ReadExcelFile(string filePath, object userData)
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
            List<int> usedLines = new List<int>();

            while ((range.Cells[row, 1] as Range).Value2 == null)
            {

                DataRow dataRow = dataTable.NewRow();

                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    dataRow[col - 1] = (range.Cells[row, col] as Range).Value2;
                }

                dataTable.Rows.Add(dataRow);

                usedLines.Add(row);
                row++;
            }

            Console.WriteLine("Used lines: " + string.Join(", ", usedLines));

            workbook.Close(true);
            excelApp.Quit();

            return dataTable;
        }

        private void OpenPart2_Click(object sender, RoutedEventArgs e)
        {
            //show the "open user management console" button
            if (UsersTextBox.Text == "Admin")
            {
                part2 newWindow = new part2();
                newWindow.Show();
            }
            else
            {
                MessageBox.Show("You must be an Admin to open this window.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Users(object sender, TextChangedEventArgs e)
        {

        }
    }
}