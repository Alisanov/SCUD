using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace SCUD
{
    public partial class MainWindow : System.Windows.Window
    {
        private readonly Connection connect;
        private readonly List<string> errors = new List<string>();
        //для хранения id справочников
        private List<int> idWorkarea = new List<int>();
        private List<int> idDepartament = new List<int>();
        private List<int> idPosition = new List<int>();

        Employee employeeActual = new Employee();

        //настройки
        readonly double animationSpeed = 0.5;

        public MainWindow()
        {
            InitializeComponent();
            connect = new Connection();
            loginWindow.Visibility = Visibility.Visible;
            panel.Visibility = Visibility.Collapsed;
            directoresArea.Visibility = Visibility.Collapsed;
            employeeArea.Visibility = Visibility.Collapsed;
            panel.IsEnabled = false;
            directoresArea.IsEnabled = false;
            employeeArea.IsEnabled = false;
            dateEmployee.SelectedDate = DateTime.Now;
        }
        //методы области входа
        public void OpenMain()
        {
            DoubleAnimation animation = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            animation.Completed += OpenMain_Completed;
            loginWindow.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);
            loginWindow.IsEnabled = false;

        }
        private void OpenMain_Completed(object sender, EventArgs e)
        {
            SizeWindow(1000, 500);
            DoubleAnimation animation = new DoubleAnimation();
            login.Visibility = Visibility.Collapsed;
            panel.Visibility = Visibility.Visible;
            animation.From = 0;
            animation.To = 1;
            animation.Duration = TimeSpan.FromSeconds(animationSpeed);
            panel.BeginAnimation(Grid.OpacityProperty, animation);
            panel.IsEnabled = true;
            LoadPanel();

        }
        private void InfomationLabel()
        {
            DoubleAnimation animation = new DoubleAnimation()
            {
                To = 0,
                Duration = new Duration(TimeSpan.FromSeconds(0.5))
            };
            animation.Completed += InfomationLabelLater;
            infoBar.BeginAnimation(HeightProperty, animation);
        }
        private void InfomationLabelLater(object sender, EventArgs e)
        {
            DoubleAnimation animation = new DoubleAnimation()
            {
                To = 30,
                Duration = new Duration(TimeSpan.FromSeconds(0.5))
            };
            infoBar.Foreground = new System.Windows.Media.SolidColorBrush(Colors.Red);
            if (errors.Count != 0)
            {
                infoBar.Text = errors[0];
                errors.Clear();
            }
            else
            {
                infoBar.Text = "Ошибка заполнения полей!";
                errors.Clear();
            }
            infoBar.BeginAnimation(HeightProperty, animation);
        }
        private void Check()
        {
            if (login.Text.Equals(string.Empty) && password.Password.Equals(string.Empty))
            {
                errors.Add("Ошибка заполнения полей!");
            }

        }
        private void Enter(object sender, RoutedEventArgs e)
        {
            Check();
            if (errors.Count == 0)
                using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
                {
                    connection.Open();
                    SQLiteCommand command = new SQLiteCommand(@"SELECT COUNT(1),[id],[role] FROM [users] WHERE [login]=@login and [password]=@password", connection);
                    command.Parameters.AddWithValue("@login", login.Text);
                    command.Parameters.AddWithValue("@password", Encryption.GetHash(password.Password));
                    SQLiteDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        int count = Convert.ToInt32(reader[0].ToString());
                        if (count == 1)
                        {
                            if ((int.Parse(reader[2].ToString()) == 1))
                            {
                                mainText.Text = "Привет, Администратор, тебе доступны следующие действия...";
                                OpenMain();
                            }
                        }
                        else
                        {
                            errors.Add("Ошибка входа!");
                            InfomationLabel();
                        }
                    }
                }
            else
                InfomationLabel();
        }
        //конец методов области входа

        //методы главная область
        private void LoadPanel()
        {
            Employee employee = new Employee();
            employee.LoadEmployee(ref employeeListPanel);
            employee.ActualEmployee(DateTime.Now, ref employeeListActual);
            employeeActualText.Text = "Список сотрудников в организации на " + DateTime.Now.ToShortDateString();
        }
        private void EmployeeClick(object sender, RoutedEventArgs e)
        {
            DoubleAnimation animation = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            animation.Completed += EployeeClick_Completed;
            panel.BeginAnimation(OpacityProperty, animation);
            panel.IsEnabled = false;
            LoadEmployee();
        }
        private void DirictoryClick(object sender, RoutedEventArgs e)//открыть справочники
        {
            OpenDirictory();
            new Directory(ref departamentList, ref workareaList, ref positionList);
        }
        //поиск сотрудника
        private void FindEmployee()
        {
            Employee employee = new Employee();
            if (!findText.Text.Equals(string.Empty))
                employee.FindEmployee(ref employeeListPanel, findText.Text);
            else
                employee.LoadEmployee(ref employeeListPanel);
        }
        private void FindTextChanged(object sender, TextChangedEventArgs e)
        {
            FindEmployee();
        }
        private void FindClick(object sender, RoutedEventArgs e)
        {
            FindEmployee();
        }
        private void EmployeeListPanelSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (employeeListPanel.SelectedIndex != -1)
            {
                employeeActual = (Employee)employeeListPanel.SelectedItem;
                name.Content = employeeActual.Lastname + ' ' + employeeActual.Name;
            }
        }
        private void CheckInputClick(object sender, RoutedEventArgs e)
        {
            if (checkInput.IsChecked.Value)
                timeInput.IsEnabled = false;
            else
                timeInput.IsEnabled = true;
        }
        private void DateEmployeeSelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!dateEmployee.Text.Equals(string.Empty))
            {
                new Employee().ActualEmployee(dateEmployee.SelectedDate.Value, ref employeeListActual);
                employeeActualText.Text = "Список сотрудников в организации на " + dateEmployee.SelectedDate.Value.ToShortDateString();
            }
        }
        private void EmployeeGo(object sender, RoutedEventArgs e)
        {
            if (employeeListPanel.SelectedIndex != -1 || employeeListActual.SelectedIndex != -1)
            {

                if (!checkInput.IsChecked.Value)
                {
                    if (!timeStart.Text.Equals(string.Empty) && !dateEmployee.Text.Equals(string.Empty))
                    {
                        employeeActual.Go(dateEmployee.SelectedDate.Value, timeStart.Text, false);
                        new Employee().ActualEmployee(dateEmployee.SelectedDate.Value, ref employeeListActual);
                    }
                    else
                        MyMessageBox.Show("Пустое поле: время прихода и/или дата!", "Отказ", MessageBoxButton.OK);
                }
                else
                {
                    employeeActual.Go(DateTime.Now, DateTime.Now.ToShortTimeString(), false);
                    new Employee().ActualEmployee(DateTime.Now, ref employeeListActual);
                }

            }
            else
                MyMessageBox.Show("Выберите сотрудника!", "Отказ", MessageBoxButton.OK);
        }
        private void EmployeeLeave(object sender, RoutedEventArgs e)
        {
            if (employeeListPanel.SelectedIndex != -1 || employeeListActual.SelectedIndex != -1)
            {

                if (!checkInput.IsChecked.Value)
                {
                    if (!timeEnd.Text.Equals(string.Empty) && !dateEmployee.Text.Equals(string.Empty))
                    {
                        employeeActual.Go(dateEmployee.SelectedDate.Value, timeEnd.Text, true);
                        new Employee().ActualEmployee(dateEmployee.SelectedDate.Value, ref employeeListActual);
                    }
                    else
                        MyMessageBox.Show("Пустое поле: время ухода и/или дата!", "Отказ", MessageBoxButton.OK);
                }
                else
                {
                    employeeActual.Go(DateTime.Now, DateTime.Now.ToShortTimeString(), true);
                    new Employee().ActualEmployee(DateTime.Now, ref employeeListActual);
                }
            }
            else
                MyMessageBox.Show("Выберите сотрудника!", "Отказ", MessageBoxButton.OK);
        }
        private void EmployeeListActualSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (employeeListActual.SelectedIndex != -1)
            {
                employeeActual = (Employee)employeeListActual.SelectedItem;
                name.Content = employeeActual.Lastname + ' ' + employeeActual.Name;
            }
        }
        private void GenerateAReport(object sender, RoutedEventArgs e)
        {
            try
            {
                if (employeeListPanel.SelectedIndex != -1)
                {
                    if (!dateEmployee.Text.Equals(string.Empty))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage excelPackage = new ExcelPackage())
                        {

                            excelPackage.Workbook.Properties.Author = "Cистема контроля и управления доступом";
                            excelPackage.Workbook.Properties.Title = employeeActual.Lastname + " " + employeeActual;
                            excelPackage.Workbook.Properties.Created = DateTime.Now;

                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                            worksheet.View.ShowGridLines = false;
                            worksheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";
                            //заголовок
                            worksheet.Cells["A1:L1"].Merge = true;
                            worksheet.Cells["A1"].Style.Font.Size = 18;
                            worksheet.Cells["A1"].Style.Font.Bold = true;
                            //worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.LightSkyBlue);
                            worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["A1"].Value = "Отчет \"Учет рабочего времени\"";

                            //worksheet.Row(2).Height = 5;
                            //worksheet.Row(3).Height = 5;

                            worksheet.Cells["H2:L2"].Merge = true;
                            worksheet.Cells["H2"].Style.Font.Size = 9;
                            worksheet.Cells["H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            worksheet.Cells["H2"].Value = "От " + dateEmployee.DisplayDate.ToString("01.MM.yyyy") + " До " + (dateEmployee.DisplayDate.AddMonths(1).AddDays(-dateEmployee.DisplayDate.Day).ToString("d"));


                            worksheet.Column(1).Width = 8;
                            worksheet.Column(4).Width = 9;
                            worksheet.Column(6).Width = 11;
                            worksheet.Column(7).Width = 6;
                            worksheet.Column(9).Width = 11;
                            worksheet.Column(10).Width = 6;
                            worksheet.Column(12).Width = 16;

                            worksheet.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["B4:E4"].Merge = true;
                            worksheet.Cells["B4"].Style.Font.Size = 10;
                            worksheet.Cells["B4"].Value = "Использовались корректировки";

                            worksheet.Cells["A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            worksheet.Cells["B6:C6"].Merge = true;
                            worksheet.Cells["B6"].Style.Font.Size = 10;
                            worksheet.Cells["B6"].Value = "Превышение нормы";

                            worksheet.Cells["D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["D6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            worksheet.Cells["E6:F6"].Merge = true;
                            worksheet.Cells["E6"].Style.Font.Size = 10;
                            worksheet.Cells["E6"].Value = "Невыполнение нормы";

                            worksheet.Cells["G6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["G6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                            worksheet.Cells["H6:I6"].Merge = true;
                            worksheet.Cells["H6"].Style.Font.Size = 10;
                            worksheet.Cells["H6"].Value = "Время предыдущего дня";

                            worksheet.Cells["J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["J6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                            worksheet.Cells["K6:L6"].Merge = true;
                            worksheet.Cells["K6"].Style.Font.Size = 10;
                            worksheet.Cells["K6"].Value = "Время следующего дня";

                            for (int i = 8; i <= 13; i++)
                            {
                                worksheet.Cells["A" + i + ":E" + i].Merge = true;
                                worksheet.Cells["F" + i + ":L" + i].Merge = true;
                            }

                            ExcelRange UsedRange = worksheet.Cells["A8:L14"];
                            UsedRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            UsedRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            UsedRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            UsedRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                            worksheet.Cells["A8:E13"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells["A8:E13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A8:E13"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            worksheet.Cells["A8:E13"].Style.Font.Bold = true;
                            worksheet.Cells["A8:E13"].Style.Font.Name = "Tahoma";
                            worksheet.Cells["A8:E13"].Style.Font.Size = (float)8.5;

                            worksheet.Cells["A8"].Value = "Сотрудник";
                            worksheet.Cells["A9"].Value = "Табельный номер";
                            worksheet.Cells["A10"].Value = "Отдел";
                            worksheet.Cells["A11"].Value = "Должность";
                            worksheet.Cells["A12"].Value = "График работы";
                            worksheet.Cells["A13"].Value = "Рабочая область";

                            worksheet.Cells["F8:L13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["F8:L13"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
                            worksheet.Cells["F8:L13"].Style.Font.Size = 10;

                            worksheet.Cells["F8"].Value = employeeActual.Lastname + " " + employeeActual.Name;
                            worksheet.Cells["F10"].Value = employeeActual.Departаment;
                            worksheet.Cells["F11"].Value = employeeActual.Position;
                            worksheet.Cells["F12"].Value = "График работы";
                            worksheet.Cells["F13"].Value = employeeActual.Workarea;

                            worksheet.Cells["G14:H14"].Merge = true;
                            worksheet.Cells["J14:K14"].Merge = true;

                            worksheet.Cells["A14:L14"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells["A14:L14"].Style.Font.Bold = true;
                            worksheet.Cells["A14:L14"].Style.Font.Name = "Tahoma";
                            worksheet.Cells["A14:L14"].Style.Font.Size = (float)8.5;
                            worksheet.Cells["A14"].Value = "Дата";
                            worksheet.Cells["B14"].Value = "Приход";
                            worksheet.Cells["C14"].Value = "Уход";
                            worksheet.Cells["D14"].Value = "Отработка";
                            worksheet.Cells["E14"].Value = "Норма";
                            worksheet.Cells["F14"].Value = "Недоработка";
                            worksheet.Cells["G14"].Value = "Переработка";
                            worksheet.Cells["I14"].Value = "Опоздание";
                            worksheet.Cells["J14"].Value = "Раний уход";
                            worksheet.Cells["L14"].Value = "Время отсутствия";

                            int days = dateEmployee.DisplayDate.AddMonths(1).AddDays(-dateEmployee.DisplayDate.Day).Day;
                            DateTime date = dateEmployee.DisplayDate.AddDays(1).AddDays(-dateEmployee.DisplayDate.Day);//первый день месяца


                            TimeSpan timeStartWorking = TimeSpan.Parse("09:00");//начала работы   
                            TimeSpan timeEndWorking = TimeSpan.Parse("16:45");//конец рабочего дня
                            TimeSpan timeNormalWorking = timeEndWorking.Subtract(timeStartWorking);
                            List<Employee> employeeListDate = employeeActual.EmployeeDateList(date);
                            //для хранения итогов
                            TimeSpan timeTotalWorkingOff = new TimeSpan();
                            TimeSpan timeTotalNormalWorking = new TimeSpan();
                            TimeSpan timeTotalDefect = new TimeSpan();
                            TimeSpan timeTotalProcessing = new TimeSpan();
                            TimeSpan timeTotalDelay = new TimeSpan();
                            TimeSpan timeTotalLeave = new TimeSpan();
                            TimeSpan timeTotalNone = new TimeSpan();
                            TimeSpan timeNone = timeNormalWorking;

                            bool flag = true;

                            for (int i = 0; i < days; i++)
                            {
                                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)//убираем выходные
                                {
                                    ExcelRange range = worksheet.Cells["A" + (15 + i) + ":L" + (15 + i)];
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells["A" + (15 + i) + ":L" + (15 + i)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells["A" + (15 + i) + ":L" + (15 + i)].Style.Font.Name = "Microsoft Sans Serif";
                                    worksheet.Cells["A" + (15 + i) + ":L" + (15 + i)].Style.Font.Size = (float)8.5;

                                    worksheet.Cells["A" + (15 + i)].Value = date.ToString("d");

                                    worksheet.Cells["G" + (15 + i) + ":H" + (15 + i)].Merge = true;
                                    worksheet.Cells["J" + (15 + i) + ":K" + (15 + i)].Merge = true;

                                    worksheet.Cells["E" + (15 + i)].Value = timeNormalWorking.ToString(@"hh\:mm");//норма, в дальнейшем можно сделать изменяемой
                                    timeTotalNormalWorking += timeNormalWorking;

                                    foreach (Employee employee in employeeListDate)
                                    {
                                        if (date.Equals(employee.Date))
                                        {
                                            flag = false;
                                            worksheet.Cells["B" + (15 + i)].Value = employee.TimeStartExcel;
                                            if (employee.TimeEndExcel != "NULL")
                                            {
                                                worksheet.Cells["C" + (15 + i)].Value = employee.TimeEndExcel;
                                                TimeSpan timeEnd = TimeSpan.Parse(employee.TimeEndExcel);
                                                TimeSpan timeStart = TimeSpan.Parse(employee.TimeStartExcel);

                                                TimeSpan timeWorkingOff = timeEnd - timeStart; //вычисление отработки'
                                                timeTotalWorkingOff += timeWorkingOff;
                                                worksheet.Cells["D" + (15 + i)].Value = timeWorkingOff.ToString(@"hh\:mm");// вывод отработки
                                                if (timeNormalWorking > timeWorkingOff)
                                                {
                                                    TimeSpan timeDefect = timeNormalWorking - timeWorkingOff;
                                                    timeTotalDefect += timeDefect;
                                                    worksheet.Cells["F" + (15 + i)].Value = timeDefect.ToString(@"hh\:mm");//вывод недоработки
                                                    worksheet.Cells["F" + (15 + i)].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                                    worksheet.Cells["G" + (15 + i)].Value = '-';
                                                }
                                                else
                                                {
                                                    TimeSpan timeProcessing = timeWorkingOff - timeNormalWorking;//вычисление переработки
                                                    timeTotalProcessing += timeProcessing;
                                                    worksheet.Cells["G" + (15 + i)].Value = timeProcessing.ToString(@"hh\:mm");
                                                    worksheet.Cells["G" + (15 + i)].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                                                    worksheet.Cells["F" + (15 + i)].Value = '-';
                                                }

                                                TimeSpan timeDelay = new TimeSpan();
                                                TimeSpan timeLeave = new TimeSpan();

                                                if (timeStartWorking < timeStart)
                                                {
                                                    timeDelay = timeStart.Subtract(timeStartWorking);//вычисления опоздании
                                                    timeTotalDelay += timeDelay;
                                                    worksheet.Cells["I" + (15 + i)].Value = timeDelay.ToString(@"hh\:mm");
                                                }
                                                else
                                                    worksheet.Cells["I" + (15 + i)].Value = '-';

                                                if (timeEndWorking > timeEnd)
                                                {
                                                    if (timeEnd < timeStartWorking)
                                                        timeLeave = timeNormalWorking;
                                                    else
                                                        timeLeave = timeEndWorking.Subtract(timeEnd);//вычисление раних уходов
                                                    timeTotalLeave += timeLeave;
                                                    worksheet.Cells["J" + (15 + i)].Value = timeLeave.ToString(@"hh\:mm");
                                                }
                                                else
                                                    worksheet.Cells["J" + (15 + i)].Value = '-';

                                                timeNone = timeDelay + timeLeave; //вычисление время отсутствия


                                            }
                                            else
                                                worksheet.Cells["C" + (15 + i)].Value = '-';
                                        }
                                    }

                                    worksheet.Cells["L" + (15 + i)].Value = timeNone.ToString(@"hh\:mm");//время отсутствия, в дальнейшем можно сделать изменяемой
                                    timeTotalNone += timeNone;
                                    timeNone = timeNormalWorking;

                                    if (flag)
                                    {
                                        worksheet.Cells["B" + (15 + i)].Value = '-';
                                        worksheet.Cells["C" + (15 + i)].Value = '-';
                                        worksheet.Cells["D" + (15 + i)].Value = '-';
                                        worksheet.Cells["F" + (15 + i)].Value = '-';
                                        worksheet.Cells["G" + (15 + i)].Value = '-';
                                        worksheet.Cells["I" + (15 + i)].Value = '-';
                                        worksheet.Cells["J" + (15 + i)].Value = '-';

                                    }
                                    else
                                        flag = true;

                                }
                                else
                                {
                                    days -= 1;
                                    i -= 1;
                                }

                                if (days == (i + 1))
                                {
                                    ExcelRange range = worksheet.Cells["A" + (16 + i) + ":L" + (16 + i)];
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    worksheet.Row(16 + i).Height = 20;
                                    worksheet.Cells["G" + (16 + i) + ":H" + (16 + i)].Merge = true;
                                    worksheet.Cells["J" + (16 + i) + ":K" + (16 + i)].Merge = true;
                                    worksheet.Cells["A" + (16 + i) + ":C" + (16 + i)].Merge = true;
                                    worksheet.Cells["A" + (16 + i)].Value = "Итого";
                                    worksheet.Cells["D" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalWorkingOff.TotalHours, timeTotalWorkingOff.Minutes);
                                    worksheet.Cells["E" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalNormalWorking.TotalHours, timeTotalNormalWorking.Minutes);
                                    worksheet.Cells["F" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalDefect.TotalHours, timeTotalDefect.Minutes);
                                    worksheet.Cells["G" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalProcessing.TotalHours, timeTotalProcessing.Minutes);
                                    worksheet.Cells["I" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalDelay.TotalHours, timeTotalDelay.Minutes);
                                    worksheet.Cells["J" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalLeave.TotalHours, timeTotalLeave.Minutes);
                                    worksheet.Cells["L" + (16 + i)].Value = string.Format("{0}:{1}", (int)timeTotalNone.TotalHours, timeTotalNone.Minutes);

                                }
                                date = date.AddDays(1);
                            }

                            DirectoryInfo directoryInfo = new DirectoryInfo(@"отчеты");
                            directoryInfo.Create();
                            FileInfo fi = new FileInfo(@"отчеты\Учет рабочего времени " + employeeActual.Lastname + ' ' + employeeActual.Name + ' ' + DateTime.Now.ToString("MMMM yyyy") + ".xlsx");

                            excelPackage.SaveAs(fi);

                            var proc = new Process
                            {
                                StartInfo = new ProcessStartInfo(fi.FullName)
                                {
                                    UseShellExecute = true
                                }
                            };
                            proc.Start();
                        }

                    }
                }
                else
                    MyMessageBox.Show("Выберите сотрудника из списка! ", "Отказ", MessageBoxButton.OK);
            }
            catch (Exception exc)
            {
                MyMessageBox.Show(exc.Message, exc.Source, MessageBoxButton.OK);
            }
        }
        private void GenerateAReportAll(object sender, RoutedEventArgs e)
        {
            try
            {
                if (employeeListPanel.Items.Count != 0)
                {
                    if (!dateEmployee.Text.Equals(string.Empty))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage excelPackage = new ExcelPackage())
                        {
                            Employee employeeСurrent = new Employee();
                            excelPackage.Workbook.Properties.Author = "Cистема контроля и управления доступом";
                            excelPackage.Workbook.Properties.Subject = " ";
                            excelPackage.Workbook.Properties.Created = DateTime.Now;

                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                            worksheet.View.ShowGridLines = false;
                            worksheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";
                            //заголовок
                            worksheet.Cells["A1:L1"].Merge = true;
                            worksheet.Cells["A1"].Style.Font.Size = 18;
                            worksheet.Cells["A1"].Style.Font.Bold = true;
                            //worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.LightSkyBlue);
                            worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells["A1"].Value = "Отчет \"Учет рабочего времени\"";

                            //worksheet.Row(2).Height = 5;
                            //worksheet.Row(3).Height = 5;

                            worksheet.Cells["H2:L2"].Merge = true;
                            worksheet.Cells["H2"].Style.Font.Size = 9;
                            worksheet.Cells["H2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            worksheet.Cells["H2"].Value = "От " + dateEmployee.DisplayDate.ToString("01.MM.yyyy") + " До " + (dateEmployee.DisplayDate.AddMonths(1).AddDays(-dateEmployee.DisplayDate.Day).ToString("d"));


                            worksheet.Column(1).Width = 8;
                            worksheet.Column(4).Width = 9;
                            worksheet.Column(6).Width = 11;
                            worksheet.Column(7).Width = 6;
                            worksheet.Column(9).Width = 11;
                            worksheet.Column(10).Width = 6;
                            worksheet.Column(12).Width = 16;

                            worksheet.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            worksheet.Cells["B4:E4"].Merge = true;
                            worksheet.Cells["B4"].Style.Font.Size = 10;
                            worksheet.Cells["B4"].Value = "Использовались корректировки";

                            worksheet.Cells["A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            worksheet.Cells["B6:C6"].Merge = true;
                            worksheet.Cells["B6"].Style.Font.Size = 10;
                            worksheet.Cells["B6"].Value = "Превышение нормы";

                            worksheet.Cells["D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["D6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            worksheet.Cells["E6:F6"].Merge = true;
                            worksheet.Cells["E6"].Style.Font.Size = 10;
                            worksheet.Cells["E6"].Value = "Невыполнение нормы";

                            worksheet.Cells["G6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["G6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                            worksheet.Cells["H6:I6"].Merge = true;
                            worksheet.Cells["H6"].Style.Font.Size = 10;
                            worksheet.Cells["H6"].Value = "Время предыдущего дня";

                            worksheet.Cells["J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["J6"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                            worksheet.Cells["K6:L6"].Merge = true;
                            worksheet.Cells["K6"].Style.Font.Size = 10;
                            worksheet.Cells["K6"].Value = "Время следующего дня";

                            int newEmployeeCells = 8;
                            int days = dateEmployee.DisplayDate.AddMonths(1).AddDays(-dateEmployee.DisplayDate.Day).Day;
                            DateTime date = dateEmployee.DisplayDate.AddDays(1).AddDays(-dateEmployee.DisplayDate.Day);//первый день месяца

                            for (int nomerEmployee = 0; nomerEmployee < employeeListPanel.Items.Count; nomerEmployee++)
                            {
                                employeeСurrent = (Employee)employeeListPanel.Items[nomerEmployee];
                                //newEmployeeCells += 9;

                                for (int j = 0; j <= 5; j++)
                                {
                                    worksheet.Cells[(newEmployeeCells + j), 1, (newEmployeeCells + j), 5].Merge = true;
                                    worksheet.Cells[(newEmployeeCells + j), 6, (newEmployeeCells + j), 12].Merge = true;
                                }

                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 5), 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 5), 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 5), 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 5), 5].Style.Font.Bold = true;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 5), 5].Style.Font.Name = "Tahoma";
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 5), 5].Style.Font.Size = (float)8.5;

                                worksheet.Cells[newEmployeeCells, 1].Value = "Сотрудник";
                                worksheet.Cells[(newEmployeeCells + 1), 1].Value = "Табельный номер";
                                worksheet.Cells[(newEmployeeCells + 2), 1].Value = "Отдел";
                                worksheet.Cells[(newEmployeeCells + 3), 1].Value = "Должность";
                                worksheet.Cells[(newEmployeeCells + 4), 1].Value = "График работы";
                                worksheet.Cells[(newEmployeeCells + 5), 1].Value = "Рабочая область";

                                worksheet.Cells[newEmployeeCells, 6, (newEmployeeCells + 5), 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[newEmployeeCells, 6, (newEmployeeCells + 5), 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
                                worksheet.Cells[newEmployeeCells, 6, (newEmployeeCells + 5), 12].Style.Font.Size = 10;

                                worksheet.Cells[newEmployeeCells, 6].Value = employeeСurrent.Lastname + " " + employeeСurrent.Name;
                                worksheet.Cells[(newEmployeeCells + 2), 6].Value = employeeСurrent.Departаment;
                                worksheet.Cells[(newEmployeeCells + 3), 6].Value = employeeСurrent.Position;
                                worksheet.Cells[(newEmployeeCells + 4), 6].Value = "График работы";
                                worksheet.Cells[(newEmployeeCells + 5), 6].Value = employeeСurrent.Workarea;

                                worksheet.Cells[(newEmployeeCells + 6), 7, (newEmployeeCells + 6), 8].Merge = true;
                                worksheet.Cells[(newEmployeeCells + 6), 10, (newEmployeeCells + 6), 11].Merge = true;

                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Font.Bold = true;
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Font.Name = "Tahoma";
                                worksheet.Cells[newEmployeeCells, 1, (newEmployeeCells + 6), 12].Style.Font.Size = (float)8.5;
                                worksheet.Cells[(newEmployeeCells + 6), 1].Value = "Дата";
                                worksheet.Cells[(newEmployeeCells + 6), 2].Value = "Приход";
                                worksheet.Cells[(newEmployeeCells + 6), 3].Value = "Уход";
                                worksheet.Cells[(newEmployeeCells + 6), 4].Value = "Отработка";
                                worksheet.Cells[(newEmployeeCells + 6), 5].Value = "Норма";
                                worksheet.Cells[(newEmployeeCells + 6), 6].Value = "Недоработка";
                                worksheet.Cells[(newEmployeeCells + 6), 7].Value = "Переработка";
                                worksheet.Cells[(newEmployeeCells + 6), 9].Value = "Опоздание";
                                worksheet.Cells[(newEmployeeCells + 6), 10].Value = "Раний уход";
                                worksheet.Cells[(newEmployeeCells + 6), 12].Value = "Время отсутствия";



                                TimeSpan timeStartWorking = TimeSpan.Parse("09:00");//начала работы   
                                TimeSpan timeEndWorking = TimeSpan.Parse("16:45");//конец рабочего дня
                                TimeSpan timeNormalWorking = timeEndWorking.Subtract(timeStartWorking);
                                List<Employee> employeeListDate = employeeСurrent.EmployeeDateList(date);
                                //для хранения итогов
                                TimeSpan timeTotalWorkingOff = new TimeSpan();
                                TimeSpan timeTotalNormalWorking = new TimeSpan();
                                TimeSpan timeTotalDefect = new TimeSpan();
                                TimeSpan timeTotalProcessing = new TimeSpan();
                                TimeSpan timeTotalDelay = new TimeSpan();
                                TimeSpan timeTotalLeave = new TimeSpan();
                                TimeSpan timeTotalNone = new TimeSpan();
                                TimeSpan timeNone = timeNormalWorking;

                                bool flag = true;
                                newEmployeeCells += 7;
                                int i = 0;
                                for (i = 0; i < days; i++)
                                {
                                    if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)//убираем выходные
                                    {
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.Font.Name = "Microsoft Sans Serif";
                                        worksheet.Cells[(newEmployeeCells + i), 1, (newEmployeeCells + i), 12].Style.Font.Size = (float)8.5;

                                        worksheet.Cells[(newEmployeeCells + i), 1].Value = date.ToString("d");

                                        worksheet.Cells[(newEmployeeCells + i), 7, (newEmployeeCells + i), 8].Merge = true;
                                        worksheet.Cells[(newEmployeeCells + i), 10, (newEmployeeCells + i), 11].Merge = true;

                                        worksheet.Cells[(newEmployeeCells + i), 5].Value = timeNormalWorking.ToString(@"hh\:mm");//норма, в дальнейшем можно сделать изменяемой
                                        timeTotalNormalWorking += timeNormalWorking;


                                        foreach (Employee employee in employeeListDate)
                                        {
                                            if (date.Equals(employee.Date))
                                            {
                                                flag = false;
                                                worksheet.Cells[(newEmployeeCells + i), 2].Value = employee.TimeStartExcel;
                                                if (employee.TimeEndExcel != "NULL")
                                                {
                                                    worksheet.Cells[(newEmployeeCells + i), 3].Value = employee.TimeEndExcel;
                                                    TimeSpan timeEnd = TimeSpan.Parse(employee.TimeEndExcel);
                                                    TimeSpan timeStart = TimeSpan.Parse(employee.TimeStartExcel);

                                                    TimeSpan timeWorkingOff = timeEnd - timeStart; //вычисление отработки'
                                                    timeTotalWorkingOff += timeWorkingOff;
                                                    worksheet.Cells[(newEmployeeCells + i), 4].Value = timeWorkingOff.ToString(@"hh\:mm");// вывод отработки
                                                    if (timeNormalWorking > timeWorkingOff)
                                                    {
                                                        TimeSpan timeDefect = timeNormalWorking - timeWorkingOff;
                                                        timeTotalDefect += timeDefect;
                                                        worksheet.Cells[(newEmployeeCells + i), 6].Value = timeDefect.ToString(@"hh\:mm");//вывод недоработки
                                                        worksheet.Cells[(newEmployeeCells + i), 6].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                                        worksheet.Cells[(newEmployeeCells + i), 7].Value = '-';
                                                    }
                                                    else
                                                    {
                                                        TimeSpan timeProcessing = timeWorkingOff - timeNormalWorking;//вычисление переработки
                                                        timeTotalProcessing += timeProcessing;
                                                        worksheet.Cells[(newEmployeeCells + i), 7].Value = timeProcessing.ToString(@"hh\:mm");
                                                        worksheet.Cells[(newEmployeeCells + i), 7].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                                                        worksheet.Cells[(newEmployeeCells + i), 6].Value = '-';
                                                    }

                                                    TimeSpan timeDelay = new TimeSpan();
                                                    TimeSpan timeLeave = new TimeSpan();

                                                    if (timeStartWorking < timeStart)
                                                    {
                                                        timeDelay = timeStart.Subtract(timeStartWorking);//вычисления опоздании
                                                        timeTotalDelay += timeDelay;
                                                        worksheet.Cells[(newEmployeeCells + i), 9].Value = timeDelay.ToString(@"hh\:mm");
                                                    }
                                                    else
                                                        worksheet.Cells[(newEmployeeCells + i), 9].Value = '-';

                                                    if (timeEndWorking > timeEnd)
                                                    {
                                                        if (timeEnd < timeStartWorking)
                                                            timeLeave = timeNormalWorking;
                                                        else
                                                            timeLeave = timeEndWorking.Subtract(timeEnd);//вычисление раних уходов
                                                        timeTotalLeave += timeLeave;
                                                        worksheet.Cells[(newEmployeeCells + i), 10].Value = timeLeave.ToString(@"hh\:mm");
                                                    }
                                                    else
                                                        worksheet.Cells[(newEmployeeCells + i), 10].Value = '-';

                                                    timeNone = timeDelay + timeLeave; //вычисление время отсутствия


                                                }
                                                else
                                                    worksheet.Cells[(newEmployeeCells + i), 3].Value = '-';
                                            }
                                        }

                                        worksheet.Cells[(newEmployeeCells + i), 12].Value = timeNone.ToString(@"hh\:mm");//время отсутствия, в дальнейшем можно сделать изменяемой
                                        timeTotalNone += timeNone;
                                        timeNone = timeNormalWorking;

                                        if (flag)
                                        {
                                            worksheet.Cells[(newEmployeeCells + i), 2].Value = '-';
                                            worksheet.Cells[(newEmployeeCells + i), 3].Value = '-';
                                            worksheet.Cells[(newEmployeeCells + i), 4].Value = '-';
                                            worksheet.Cells[(newEmployeeCells + i), 6].Value = '-';
                                            worksheet.Cells[(newEmployeeCells + i), 7].Value = '-';
                                            worksheet.Cells[(newEmployeeCells + i), 9].Value = '-';
                                            worksheet.Cells[(newEmployeeCells + i), 10].Value = '-';

                                        }
                                        else
                                            flag = true;

                                    }
                                    else
                                    {
                                        days -= 1;
                                        i -= 1;
                                    }
                                    if (days == (i + 1))
                                    {
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 12].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet.Row(newEmployeeCells + 1 + i).Height = 20;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 7, (newEmployeeCells + 1 + i), 8].Merge = true;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 10, (newEmployeeCells + 1 + i), 11].Merge = true;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1, (newEmployeeCells + 1 + i), 3].Merge = true;
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 1].Value = "Итого";
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 4].Value = string.Format("{0}:{1}", (int)timeTotalWorkingOff.TotalHours, timeTotalWorkingOff.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 5].Value = string.Format("{0}:{1}", (int)timeTotalNormalWorking.TotalHours, timeTotalNormalWorking.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 6].Value = string.Format("{0}:{1}", (int)timeTotalDefect.TotalHours, timeTotalDefect.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 7].Value = string.Format("{0}:{1}", (int)timeTotalProcessing.TotalHours, timeTotalProcessing.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 9].Value = string.Format("{0}:{1}", (int)timeTotalDelay.TotalHours, timeTotalDelay.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 10].Value = string.Format("{0}:{1}", (int)timeTotalLeave.TotalHours, timeTotalLeave.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 1 + i), 12].Value = string.Format("{0}:{1}", (int)timeTotalNone.TotalHours, timeTotalNone.Minutes);
                                        worksheet.Cells[(newEmployeeCells + 2 + i), 6].Value = string.Empty;
                                    }
                                    date = date.AddDays(1);
                                }
                                //newEmployeeCells -= 7;
                                days = dateEmployee.DisplayDate.AddMonths(1).AddDays(-dateEmployee.DisplayDate.Day).Day;
                                date = dateEmployee.DisplayDate.AddDays(1).AddDays(-dateEmployee.DisplayDate.Day);
                                newEmployeeCells += (i + 1);

                            }
                            DirectoryInfo directoryInfo = new DirectoryInfo(@"отчеты");
                            directoryInfo.Create();
                            FileInfo fi = new FileInfo(@"отчеты\Учет рабочего времени " + dateEmployee.SelectedDate.Value.ToString("MMMM yyyy") + ".xlsx");

                            excelPackage.SaveAs(fi);

                            var proc = new Process
                            {
                                StartInfo = new ProcessStartInfo(fi.FullName)
                                {
                                    UseShellExecute = true
                                }
                            };
                            proc.Start();
                        }
                    }
                    else
                        MyMessageBox.Show("Выберите дату", "Отказ", MessageBoxButton.OK);
                }
                else
                    MyMessageBox.Show("В базе отсутствуют сотрудники ", "Отказ ", MessageBoxButton.OK);
            }
            catch (Exception exc)
            {
                MyMessageBox.Show(exc.Message, exc.Source, MessageBoxButton.OK);
            }
        }
        private void SelectFile(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFile = new OpenFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    InitialDirectory = @"отчеты"
            };
                if (openFile.ShowDialog().Value)
                {
                    FileInfo fi = new FileInfo(openFile.FileName);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage(fi))
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Sheet 1"];
                        List<Employee> employees = new List<Employee>();
                        int actualCells = 8;
                        string name = worksheet.Cells[actualCells, 6].Text;//имя и фамилия сотрудника

                        while (!name.Equals(string.Empty))
                        {

                            string departament = worksheet.Cells[actualCells + 2, 6].Text;//отдел сотрудника
                            string position = worksheet.Cells[actualCells + 3, 6].Text;//должность сотрудника
                            List<DateTime> delaysDate = new List<DateTime>();//дни опоздании
                            TimeSpan totalWorkingOff = new TimeSpan();//общее кол-во отработки
                            TimeSpan totalNormal = new TimeSpan();
                            DateTime firstDate = new DateTime();//первая рабочая дата месяца
                            DateTime endDate = new DateTime();//крайняя рабочая дата месяца
                            List<TimeSpan> delays = new List<TimeSpan>();//опоздания

                            DateTime dateTime = DateTime.Parse(worksheet.Cells[actualCells + 7, 1].Text);
                            firstDate = dateTime;
                            DateTime endDateTime = dateTime.AddMonths(1).AddDays(-dateTime.Day);
                            int countSCUDNormal = 0;
                            int countSCUDFact = 0;
                            if (endDateTime.DayOfWeek == DayOfWeek.Sunday)
                                endDateTime = endDateTime.AddDays(-2);
                            else if (endDateTime.DayOfWeek == DayOfWeek.Saturday)
                                endDateTime = endDateTime.AddDays(-1);
                            endDate = endDateTime;
                            for (int i = 0; dateTime != endDateTime; i++)
                            {
                                string delay = worksheet.Cells[actualCells + 7 + i, 9].Text;
                                if (!delay.Equals("-"))
                                {
                                    delays.Add(TimeSpan.Parse(delay));
                                    delaysDate.Add(dateTime);
                                }
                                string totalWorkingOffString = worksheet.Cells[actualCells + 7 + i, 4].Text;
                                if (!totalWorkingOffString.Equals("-"))
                                {
                                    totalWorkingOff += TimeSpan.Parse(totalWorkingOffString);
                                    countSCUDFact++;
                                }
                                string totalNormalString = worksheet.Cells[actualCells + 7 + i, 5].Text;
                                if (!totalNormalString.Equals("-"))
                                    totalNormal += TimeSpan.Parse(totalNormalString);
                                countSCUDNormal++;
                                dateTime = DateTime.Parse(worksheet.Cells[actualCells + 7 + i, 1].Text);
                            }
                            actualCells += countSCUDNormal + 7;

                            employees.Add(new Employee(name, departament, position, totalWorkingOff, totalNormal, countSCUDNormal, countSCUDFact, firstDate, endDate, delaysDate, delays));
                            actualCells += 1;
                            name = worksheet.Cells[actualCells, 6].Text;

                        }
                        CreateReport(employees);

                    }
                }
            }
            catch (Exception exc)
            {
                MyMessageBox.Show(exc.Message, exc.Source, MessageBoxButton.OK);
            }
        }
        private void CreateReport(List<Employee> employees)
        {
            try
            {
                if (employees.Count != 0)
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        excelPackage.Workbook.Properties.Author = "Cистема контроля и управления доступом";
                        excelPackage.Workbook.Properties.Created = DateTime.Now;

                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                        //worksheet.View.ShowGridLines = false;
                        worksheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";
                        worksheet.Cells["A1:W99"].Style.Font.Size = 12;
                        worksheet.Cells[4, 1, 4, 10].Merge = true;
                        worksheet.Cells[5, 1, 5, 10].Merge = true;
                        worksheet.Cells[6, 1, 6, 10].Merge = true;
                        worksheet.Cells[7, 1, 7, 10].Merge = true;
                        worksheet.Cells[8, 1, 8, 10].Merge = true;

                        worksheet.Column(1).Width = 8.30;
                        worksheet.Column(2).Width = 25;
                        worksheet.Column(3).Width = 16;
                        worksheet.Column(4).Width = 21;
                        worksheet.Column(5).Width = 14.30;
                        worksheet.Column(6).Width = 14;
                        worksheet.Column(7).Width = 17.30;
                        worksheet.Column(8).Width = 20.30;
                        worksheet.Column(9).Width = 45;
                        worksheet.Column(10).Width = 22;

                        worksheet.Cells[4, 1, 8, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        worksheet.Cells[9, 1, 9, 10].Merge = true;
                        worksheet.Cells[9, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        worksheet.Cells[4, 1].Value = "Приложение №2";
                        worksheet.Cells[4, 1].Style.Font.Bold = true;
                        worksheet.Cells[5, 1].Value = "К приказу №22 от 30.03.2018 г.";
                        worksheet.Cells[5, 1].Style.Font.Bold = true;
                        worksheet.Cells[6, 1].Value = "«О изменении режима рабочего времени и";
                        worksheet.Cells[7, 1].Value = "применении взысканий за нарушение";
                        worksheet.Cells[8, 1].Value = "трудовой дисциплины ООО «Прогресс»";

                        worksheet.Cells[9, 1].Value = "Табель учета рабочего времени с " + employees[0].FirstDate.ToString("dd") + " по " + employees[0].EndDate.ToString("dd MMMM");
                        worksheet.Cells[9, 1].Style.Font.Bold = true;

                        worksheet.Cells[11, 1, 12, 1].Merge = true;
                        worksheet.Cells[11, 2, 12, 2].Merge = true;
                        worksheet.Cells[11, 3, 12, 3].Merge = true;
                        worksheet.Cells[11, 4, 12, 4].Merge = true;

                        worksheet.Cells[11, 6, 11, 7].Merge = true;

                        worksheet.Cells[11, 9, 12, 9].Merge = true;
                        worksheet.Cells[11, 10, 12, 10].Merge = true;


                        ExcelRange usedRange = worksheet.Cells[11, 1, 12, 10];
                        usedRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        usedRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        usedRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        usedRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        usedRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        usedRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        usedRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        usedRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        usedRange.Style.Font.Bold = true;
                        usedRange.Style.WrapText = true;

                        worksheet.Cells[11, 1].Value = "№ п/п";
                        worksheet.Cells[11, 2].Value = "Подразделение";
                        worksheet.Cells[11, 3].Value = "ФИО";
                        worksheet.Cells[11, 4].Value = "Должность";

                        worksheet.Cells[11, 5].Value = "Штатное количество";
                        worksheet.Cells[12, 5].Value = "дней";

                        worksheet.Cells[11, 6].Value = "Фактическое количество по СКУД";
                        worksheet.Cells[12, 6].Value = "дней";
                        worksheet.Cells[12, 7].Value = "часов";

                        worksheet.Cells[11, 8].Value = "Дельта между СКУД и штатными показателями";
                        worksheet.Cells[12, 8].Value = "дней";

                        worksheet.Cells[11, 9].Value = "Примечание";
                        worksheet.Cells[11, 10].Value = "Заключение руководителя по выплате";
                        bool working = false;
                        for (int i = 0; i < employees.Count; i++)
                        {
                            Employee employee = employees[i];
                            usedRange = worksheet.Cells[13 + i, 1, 13 + i, 10];
                            usedRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            usedRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            usedRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            usedRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            usedRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Row(13 + i).Height = 60;
                            usedRange.Style.WrapText = true;

                            worksheet.Cells[13 + i, 1].Value = i + 1;
                            worksheet.Cells[13 + i, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[13 + i, 2].Value = employee.Departаment;
                            worksheet.Cells[13 + i, 3].Value = employee.LastnameFull;
                            worksheet.Cells[13 + i, 4].Value = employee.Position;
                            worksheet.Cells[13 + i, 5].Value = employee.DaySCUDNormal;
                            worksheet.Cells[13 + i, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[13 + i, 6].Value = employee.DaySCUDFact;
                            worksheet.Cells[13 + i, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[13 + i, 7].Value = string.Format("{0:0.00}", employee.TotalWorkingOff.TotalHours);
                            worksheet.Cells[13 + i, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            if (employee.TotalNormalWorking > employee.TotalWorkingOff)
                            {
                                TimeSpan span = employee.TotalNormalWorking - employee.TotalWorkingOff;
                                worksheet.Cells[13 + i, 8].Value = string.Format("{0}:{1}:{2}", (int)span.TotalDays, span.Hours, span.Minutes) + " (недоработка)";
                                working = true;
                            }
                            else if (employee.TotalNormalWorking == employee.TotalWorkingOff)
                            {
                                worksheet.Cells[13 + i, 8].Value = "00:00:00 (норма)";
                            }
                            else
                            {
                                TimeSpan span = employee.TotalWorkingOff - employee.TotalNormalWorking;
                                worksheet.Cells[13 + i, 8].Value = string.Format("{0}:{1}:{2}", (int)span.TotalDays, span.Hours, span.Minutes) + " (переработка)";
                                working = false;
                            }

                            List<DateTime> delayDates = employee.DelaysDate;
                            List<TimeSpan> delays = employee.Delays;

                            for (int count = 0; count < delays.Count; count++)
                            {
                                worksheet.Cells[13 + i, 9].Value += delayDates[count].ToString("d") + string.Format(" {0}:{1} ", (int)delays[count].TotalHours, delays[count].Minutes) + "опоздание. " + "\r\n";
                                worksheet.Row(13 + i).Height += 10;
                            }
                            if (delays.Count != 0 && working)
                            {
                                worksheet.Cells[13 + i, 10].Value = "Штраф за " + delays.Count + " опоздание/и и за недоработку.";
                            }
                            else if (delays.Count != 0 || working)
                            {
                                if (working)
                                    worksheet.Cells[13 + i, 10].Value = "Штраф за недоработку.";
                                else
                                    worksheet.Cells[13 + i, 10].Value = "Штраф за " + delays.Count + " опоздание/и";
                            }
                            else
                            {
                                worksheet.Cells[13 + i, 10].Value = " Без штрафа.";
                            }


                        }

                        DirectoryInfo directoryInfo = new DirectoryInfo(@"отчеты");
                        directoryInfo.Create();
                        FileInfo fi = new FileInfo(@"отчеты\Табель учета рабочего времени " + employees[0].FirstDate.Date.ToString("MMMM yyyy") + ".xlsx");

                        excelPackage.SaveAs(fi);

                        var proc = new Process
                        {
                            StartInfo = new ProcessStartInfo(fi.FullName)
                            {
                                UseShellExecute = true
                            }
                        };
                        proc.Start();
                    }
                }
                else
                {
                    MyMessageBox.Show("Ни один сотрудник не был найден!", "Выберите другой файл", MessageBoxButton.OK);
                }
            }
            catch (Exception exc)
            {
                MyMessageBox.Show(exc.Message, exc.Source, MessageBoxButton.OK);
            }
        }
        //конец методов главной области

        //методы области работы со сотрудниками
        private void EployeeClick_Completed(object sender, EventArgs e)
        {
            SizeWindow(1000, 600);
            DoubleAnimation animation = new DoubleAnimation();
            panel.Visibility = Visibility.Collapsed;
            employeeArea.Visibility = Visibility.Visible;
            animation.From = 0;
            animation.To = 1;
            animation.Duration = TimeSpan.FromSeconds(animationSpeed);
            employeeArea.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);

            employeeArea.IsEnabled = true;
        }
        private void LoadEmployee()
        {
            Directory directory = new Directory();
            directory.LoadComboBox(ref positionComboBox, ref idPosition, "position");
            directory.LoadComboBox(ref workareaComboBox, ref idWorkarea, "workarea");
            directory.LoadComboBox(ref departamentComboBox, ref idDepartament, "departament");
            new Employee().LoadEmployee(ref employeeList);
        }
        private void EmployeeClose(object sender, RoutedEventArgs e)
        {
            DoubleAnimation animation = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            animation.Completed += EmployeeClose_Completed;
            employeeArea.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);
            employeeArea.IsEnabled = false;
        }
        private void EmployeeClose_Completed(object sender, EventArgs e)
        {
            SizeWindow(1000, 500);

            panel.Visibility = Visibility.Visible;
            employeeArea.Visibility = Visibility.Collapsed;

            DoubleAnimation animation = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            panel.BeginAnimation(Grid.OpacityProperty, animation);
            LoadPanel();
            panel.IsEnabled = true;
        }
        private void InsertEmployeeClick(object sender, RoutedEventArgs e)
        {
            if (nameEmployee.Text == string.Empty)
                errors.Add("Имя не заполнено");
            if (lastnameEmployee.Text == string.Empty)
                errors.Add("Фамилия не заполнено");
            if (positionComboBox.SelectedIndex == -1)
                errors.Add("Выберите должность из списка");
            if (workareaComboBox.SelectedIndex == -1)
                errors.Add("Выберите рабочую область из списка");
            if (departamentComboBox.SelectedIndex == -1)
                errors.Add("Выберите отдел из списка == -1");
            if (errors.Count == 0)
            {
                Employee employee = new Employee();
                employee.InsertEmployee(nameEmployee.Text, lastnameEmployee.Text, idDepartament[departamentComboBox.SelectedIndex],
                    idWorkarea[workareaComboBox.SelectedIndex], idPosition[positionComboBox.SelectedIndex]);
                nameEmployee.Text = lastnameEmployee.Text = string.Empty;
                employee.LoadEmployee(ref employeeList);
            }
            else
                InfomationLabelEmployee();
        }
        private void DeleteEmployeeClick(object sender, RoutedEventArgs e)
        {
            if (employeeList.SelectedIndex != -1)
            {
                Employee employee = (Employee)employeeList.SelectedItem;
                employee.DeleteEmployee();
                employee.LoadEmployee(ref employeeList);
            }
            else
                MyMessageBox.Show("Выберите сотрудника из списка", "Отказ", MessageBoxButton.OK);
        }
        private void InfomationLabelEmployee()
        {
            DoubleAnimation animation = new DoubleAnimation()
            {
                To = 0,
                Duration = new Duration(TimeSpan.FromSeconds(0.5))
            };
            animation.Completed += InfomationEmployeeLabelLater;
            infoBarEmployee.BeginAnimation(HeightProperty, animation);
        }
        private void InfomationEmployeeLabelLater(object sender, EventArgs e)
        {
            DoubleAnimation animation = new DoubleAnimation()
            {
                To = 30,
                Duration = new Duration(TimeSpan.FromSeconds(0.5))
            };
            infoBarEmployee.Foreground = new System.Windows.Media.SolidColorBrush(Colors.Red);
            if (errors.Count != 0)
            {
                infoBarEmployee.Text = errors[0];
                errors.Clear();
            }
            else
            {
                infoBarEmployee.Text = "Ошибка заполнения полей!";
                errors.Clear();
            }
            infoBarEmployee.BeginAnimation(HeightProperty, animation);
        }
        //конец области работы со сотрудниками

        //методы для работы со справочниками
        public void OpenDirictory()
        {
            DoubleAnimation animation = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            animation.Completed += OpenDirictory_Completed;
            panel.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);
            panel.IsEnabled = false;
        }
        //анимация открытия
        private void OpenDirictory_Completed(object sender, EventArgs e)
        {
            SizeWindow(1000, 600);

            panel.Visibility = Visibility.Collapsed;
            directoresArea.Visibility = Visibility.Visible;

            DoubleAnimation animation = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            directoresArea.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);

            directoresArea.IsEnabled = true;
        }
        private void DeleteDirictory(object sender, RoutedEventArgs e)
        {
            if (departamentList.SelectedIndex != -1 || workareaList.SelectedIndex != -1 || positionList.SelectedIndex != -1)
            {
                if (MyMessageBox.Show("Вы уверены, что хотите удалить элемент?", "Удаление", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    Directory directory = new Directory();
                    if (departamentList.SelectedIndex != -1)
                        directory.DeleteDirectory("departament", ref departamentList);
                    if (workareaList.SelectedIndex != -1)
                        directory.DeleteDirectory("workarea", ref workareaList);
                    if (positionList.SelectedIndex != -1)
                        directory.DeleteDirectory("position", ref positionList);
                }
            }
            else
                MyMessageBox.Show("Выберите элемент из списка", "Отказ", MessageBoxButton.OK);
        }
        //справочник отделы
        private void SaveDepartment(object sender, RoutedEventArgs e)
        {
            new Directory().SaveDirectory("departament", departamentText.Text, ref departamentList);
            departamentText.Text = string.Empty;
        }
        //cправочник область работы
        private void SaveWorkarea(object sender, RoutedEventArgs e)
        {
            new Directory().SaveDirectory("workarea", workareaText.Text, ref workareaList);
            workareaText.Text = string.Empty;
        }
        //справочник должностей
        private void SavePosition(object sender, RoutedEventArgs e)
        {
            new Directory().SaveDirectory("position", positionText.Text, ref positionList);
            positionText.Text = string.Empty;
        }
        //обработчики фокуса справочников
        private void Departament_GotFocus(object sender, RoutedEventArgs e)
        {
            workareaList.SelectedIndex = -1;
            positionList.SelectedIndex = -1;
        }
        private void Workarea_GotFocus(object sender, RoutedEventArgs e)
        {
            departamentList.SelectedIndex = -1;
            positionList.SelectedIndex = -1;
        }
        private void Position_GotFocus(object sender, RoutedEventArgs e)
        {
            departamentList.SelectedIndex = -1;
            workareaList.SelectedIndex = -1;
        }
        private void CloseDirictiry(object sender, RoutedEventArgs e)
        {
            DoubleAnimation animation = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            animation.Completed += CloseDirictory_Completed;
            directoresArea.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);
            directoresArea.IsEnabled = false;
        }
        //анимация закрытия
        private void CloseDirictory_Completed(object sender, EventArgs e)
        {
            SizeWindow(1000, 500);

            panel.Visibility = Visibility.Visible;
            directoresArea.Visibility = Visibility.Collapsed;

            DoubleAnimation animation = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            panel.BeginAnimation(System.Windows.Controls.Grid.OpacityProperty, animation);
            panel.IsEnabled = true;
        }
        //конец методов работы со справочниками

        //анимация окна
        private void SizeWindow(double width, double height)
        {
            DoubleAnimation animationWidth = new DoubleAnimation
            {
                To = width,
                Duration = TimeSpan.FromSeconds(animationSpeed)
            };
            {

            }
            window.BeginAnimation(System.Windows.Controls.Grid.WidthProperty, animationWidth);

            DoubleAnimation animationHeight = new DoubleAnimation
            {
                To = height,
                Duration = TimeSpan.FromSeconds(animationSpeed + 0.5)
            };
            window.BeginAnimation(System.Windows.Controls.Grid.HeightProperty, animationHeight);
        }

    }
}
