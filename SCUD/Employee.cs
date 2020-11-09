using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SQLite;
using System.Windows;
using System.Windows.Controls;

namespace SCUD
{
    class Employee
    {
        private readonly int id;
        private readonly string name;
        private readonly string lastname;
        ///private int positionID;
        private readonly string position;
        //private int workareaID;
        private readonly string workarea;
        //private int departamentID;
        private readonly string departament;
        private readonly string timeStart;
        private readonly string timeEnd;
        private readonly DateTime date;

        //для отчета
        private string lastnameFull;
        private List<DateTime> delayDates = new List<DateTime>();
        private List<TimeSpan> delays = new List<TimeSpan>();
        private TimeSpan totalWorkingOff = new TimeSpan();
        private TimeSpan totalNormalWorking = new TimeSpan();
        private DateTime firstDate = new DateTime();
        private DateTime endDate = new DateTime();
        private int daySCUDNormal = 0;
        private int daySCUDFact = 0;

        public int ID { get { return id; } }
        public string Name { get { return name; } }
        public string Lastname { get { return lastname; } }
        public string Position { get { return position; } }
        public string Workarea { get { return workarea; } }
        public string Departаment { get { return departament; } }
        public string TimeStartExcel { get { return timeStart; } }
        public string TimeEndExcel { get { return timeEnd; } }
        public string TimeStart { get { return "Время прихода: " + timeStart; } }
        public string TimeEnd { get { return (timeEnd.Equals("NULL") ? "Сотрудник еще на работе" : "Время ухода: " + timeEnd); } }
        public DateTime Date { get { return date; } }

        public string LastnameFull { get { return lastnameFull; } }
        public List<TimeSpan> Delays { get { return delays; } set { delays = value; } }
        public List<DateTime> DelaysDate { get { return delayDates; } set { delayDates = value; } }
        public TimeSpan TotalWorkingOff { get { return totalWorkingOff; } set { totalWorkingOff = value; } }
        public TimeSpan TotalNormalWorking { get { return totalNormalWorking; } set { totalNormalWorking = value; } }
        public DateTime FirstDate { get { return firstDate; } set { firstDate = value; } }
        public DateTime EndDate { get { return endDate; } set { endDate = value; } }
        public int DaySCUDNormal { get { return daySCUDNormal; } set { daySCUDNormal = value; } }
        public int DaySCUDFact { get { return daySCUDFact; } set { daySCUDFact = value; } }
        public Employee()
        {
            id = -1;
            name = string.Empty;
            lastname = string.Empty;
            position = string.Empty;
            workarea = string.Empty;
            departament = string.Empty;
            timeStart = string.Empty;
            timeEnd = string.Empty;
        }
        public Employee(int id, string name, string lastname)
        {
            this.id = id;
            this.name = name;
            this.lastname = lastname;
            timeStart = string.Empty;
            timeEnd = string.Empty;
        }
        public Employee(int id, string name, string lastname, string timeStart, string timeEnd)
        {
            this.id = id;
            this.name = name;
            this.lastname = lastname;
            this.timeStart = timeStart;
            this.timeEnd = timeEnd;
        }
        public Employee(int id, string name, string lastname, string timeStart, string timeEnd, DateTime date)
        {
            this.id = id;
            this.name = name;
            this.lastname = lastname;
            this.timeStart = timeStart;
            this.timeEnd = timeEnd;
            this.date = date;
        }
        public Employee(int id, string name, string lastname, string position, string workarea, string departament)
        {
            this.id = id;
            this.name = name;
            this.lastname = lastname;
            this.position = position;
            this.workarea = workarea;
            this.departament = departament;
        }
        public Employee(string lastnameFull, string departament, string position, TimeSpan totalWorkingOff, TimeSpan totalNormalWorking, int daySCUD, int daySCUDFact, DateTime firstDate, DateTime endDate, List<DateTime> delayDate, List<TimeSpan> delay)
        {
            this.lastnameFull = lastnameFull;
            this.departament = departament;
            this.position = position;
            this.totalWorkingOff = totalWorkingOff;
            this.totalNormalWorking = totalNormalWorking;
            this.daySCUDNormal = daySCUD;
            this.firstDate = firstDate;
            this.endDate = endDate;
            this.delays = delay;
            this.delayDates = delayDate;
            this.daySCUDFact = daySCUDFact;
        }

        public void InsertEmployee(string name, string lastname, int idDepartament, int idWorkarea, int idPosition)
        {
            Connection connect = new Connection();
            using (SQLiteConnection sQLiteConnection = new SQLiteConnection(connect.Connect))
            {
                sQLiteConnection.Open();
                SQLiteCommand sQLiteCommand = new SQLiteCommand("INSERT INTO [employee] " +
                    "(name, lastname, departamentID, positionID, workareaID) " +
                    "VALUES(@name,@lastname,@departamentID,@positionID, @workareaID)", sQLiteConnection);
                sQLiteCommand.Parameters.AddWithValue("@name", name);
                sQLiteCommand.Parameters.AddWithValue("@lastname", lastname);
                sQLiteCommand.Parameters.AddWithValue("@departamentID", idDepartament);
                sQLiteCommand.Parameters.AddWithValue("@positionID", idPosition);
                sQLiteCommand.Parameters.AddWithValue("@workareaID", idWorkarea);
                sQLiteCommand.ExecuteNonQuery();
            }
        }
        public void LoadEmployeeLite(ref ListBox listBox)
        {
            ObservableCollection<Employee> listEmployee = new ObservableCollection<Employee>();
            Connection connect = new Connection();
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(
                    "SELECT [id], [name], [lastname] " +
                    "FROM [employee] ORDER BY lastname", connection)
                {
                    CommandType = CommandType.Text
                };
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    listEmployee.Add(new Employee(int.Parse(reader[0].ToString()), reader[1].ToString(), reader[2].ToString()));
                }
            }
            listBox.ItemsSource = listEmployee;
        }
        public void LoadEmployee(ref ListBox listBox)
        {
            ObservableCollection<Employee> listEmployee = new ObservableCollection<Employee>();
            Connection connect = new Connection();
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(
                    "SELECT employee.id, employee.name, employee.lastname, position.name, workarea.name, departament.name " +
                    "FROM [employee], [position], [departament], [workarea] " +
                    "WHERE employee.positionID = position.id AND employee.workareaID = workarea.id AND employee.departamentID = departament.id  ORDER BY lastname", connection)
                {
                    CommandType = CommandType.Text
                };
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    listEmployee.Add(new Employee(int.Parse(reader[0].ToString()), reader[1].ToString(), reader[2].ToString(), reader[3].ToString(), reader[4].ToString(), reader[5].ToString()));
                }
            }
            listBox.ItemsSource = listEmployee;
        }
        public void FindEmployee(ref ListBox listBox, string findText)
        {
            try
            {
                List<string> texts = new List<string>();

                for (int i = 0; i < findText.Length; i++)
                {
                    if (findText[i].Equals(' ') || (i + 1) == findText.Length)
                    {
                        texts.Add(findText.Substring(0, i + 1));
                        findText = findText.Remove(0, i + 1);
                        i = 0;
                    }
                }
                string commandText = string.Empty;
                foreach (string text in texts)
                {
                    string textTrim = text.Trim(' ');
                    if (!text.Equals(string.Empty))

                        commandText += " [employee].[name] LIKE \"" + textTrim + "%\" OR [employee].[lastname] LIKE \"" + textTrim + "%\" OR ";
                }
                commandText += commandText.Remove(commandText.Length - 3);
                ObservableCollection<Employee> listEmployee = new ObservableCollection<Employee>();
                Connection connect = new Connection();
                using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
                {
                    connection.Open();
                    SQLiteCommand command = new SQLiteCommand(
                         "SELECT [employee].[id], [employee].[name], [employee].[lastname], [position].[name], [workarea].[name], [departament].[name] " +
                        "FROM [employee], [position], [departament], [workarea] " +
                        "WHERE (" + commandText + ") AND [employee].[positionID] = [position].[id] AND [employee].[workareaID] = [workarea].[id] AND [employee].[departamentID] = [departament].[id]  ORDER BY [employee].[lastname]", connection)
                    {
                        CommandType = CommandType.Text
                    };
                    SQLiteDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        listEmployee.Add(new Employee(int.Parse(reader[0].ToString()), reader[1].ToString(), reader[2].ToString(), reader[3].ToString(), reader[4].ToString(), reader[5].ToString()));
                    }
                }
                listBox.ItemsSource = listEmployee;
            }
            catch (Exception err)
            {
                MyMessageBox.Show(err.Message, err.Source, MessageBoxButton.OK);
            }
        }
        public void DeleteEmployee()
        {
            if (MyMessageBox.Show("Вы уверены, что хотите удалить сотрудника?", "Удаление", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Connection connect = new Connection();
                using (SQLiteConnection sQLiteConnection = new SQLiteConnection(connect.Connect))
                {
                    sQLiteConnection.Open();
                    SQLiteCommand sQLiteCommand = new SQLiteCommand("DELETE FROM [employee] WHERE [id]=@id", sQLiteConnection);
                    sQLiteCommand.Parameters.AddWithValue("@id", id);
                    sQLiteCommand.ExecuteNonQuery();
                }
            }
        }
        public void Go(DateTime date, string time, bool leave)
        {
            Connection connection = new Connection();
            using (SQLiteConnection sQLiteConnection = new SQLiteConnection(connection.Connect))
            {
                sQLiteConnection.Open();
                SQLiteCommand sQLiteCommand = new SQLiteCommand("SELECT id FROM schedule " +
                    "WHERE employeeID=@id AND data=@date", sQLiteConnection);
                sQLiteCommand.Parameters.AddWithValue("@id", id);
                sQLiteCommand.Parameters.AddWithValue("@date", date.ToShortDateString());
                SQLiteDataReader sQLiteDataReader = sQLiteCommand.ExecuteReader();
                if (sQLiteDataReader.Read())
                {
                    sQLiteDataReader.Close();
                    sQLiteConnection.Close();
                    if (!leave)
                        UpdateScheldule(date, time);
                    else
                    {
                        UpdateSchelduleLeave(date, time);
                        return;
                    }
                }
                else
                {
                    sQLiteDataReader.Close();
                    sQLiteConnection.Close();
                    if (!leave)
                        InsertScheldule(date, time);
                    else
                        MyMessageBox.Show("Этот сотрудник даже еще не пришел сегодня в организацию!", "Отказ", MessageBoxButton.OK);
                }
            }
        }
        private void InsertScheldule(DateTime date, string time)
        {
            Connection connection = new Connection();
            using (SQLiteConnection sQLiteConnection = new SQLiteConnection(connection.Connect))
            {
                sQLiteConnection.Open();
                SQLiteCommand sQLiteCommand = new SQLiteCommand("INSERT INTO schedule (employeeID, data, timeStart, timeEnd) VALUES( @id, @date, @timeStart, @timeEnd)", sQLiteConnection);

                sQLiteCommand.Parameters.AddWithValue("@id", id);
                sQLiteCommand.Parameters.AddWithValue("@date", date.ToShortDateString());
                sQLiteCommand.Parameters.AddWithValue("@timeStart", time);
                sQLiteCommand.Parameters.AddWithValue("@timeEnd", "NULL");

                sQLiteCommand.ExecuteNonQuery();
            }
        }
        private void UpdateScheldule(DateTime date, string time)
        {
            Connection connection = new Connection();
            using (SQLiteConnection sQLiteConnection = new SQLiteConnection(connection.Connect))
            {
                sQLiteConnection.Open();
                SQLiteCommand sQLiteCommand = new SQLiteCommand("UPDATE schedule SET timeStart=@time WHERE employeeID=@id AND data=@date", sQLiteConnection);

                sQLiteCommand.Parameters.AddWithValue("@id", id);
                sQLiteCommand.Parameters.AddWithValue("@date", date.ToShortDateString());
                sQLiteCommand.Parameters.AddWithValue("@time", time);

                sQLiteCommand.ExecuteNonQuery();
            }
        }
        private void UpdateSchelduleLeave(DateTime date, string time)
        {
            Connection connection = new Connection();
            using (SQLiteConnection sQLiteConnection = new SQLiteConnection(connection.Connect))
            {
                sQLiteConnection.Open();
                SQLiteCommand sQLiteCommand = new SQLiteCommand("UPDATE schedule SET timeEnd=@time WHERE employeeID=@id AND data=@date", sQLiteConnection);

                sQLiteCommand.Parameters.AddWithValue("@id", id);
                sQLiteCommand.Parameters.AddWithValue("@date", date.ToShortDateString());
                sQLiteCommand.Parameters.AddWithValue("@time", time);

                sQLiteCommand.ExecuteNonQuery();
            }
        }
        public void ActualEmployee(DateTime date, ref ListBox listBox)
        {
            ObservableCollection<Employee> listEmployee = new ObservableCollection<Employee>();
            Connection connect = new Connection();
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(
                    "SELECT [employee].[id], [employee].[name], [employee].[lastname], [schedule].[timeStart], [schedule].[timeEnd]" +
                    "FROM [employee], [schedule] " +
                    "WHERE [schedule].[employeeID]=[employee].[id] AND [schedule].[data]=@date ", connection)
                {
                    CommandType = CommandType.Text
                };
                command.Parameters.AddWithValue("@date", date.ToShortDateString());
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    listEmployee.Add(new Employee(int.Parse(reader[0].ToString()), reader[1].ToString(), reader[2].ToString(), reader[3].ToString(), reader[4].ToString()));
                }
            }
            listBox.ItemsSource = listEmployee;
        }

        public List<Employee> EmployeeDateList(DateTime date)
        {
            List<Employee> listBox = new List<Employee>();
            Connection connect = new Connection();
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(
                    "SELECT [employee].[id], [employee].[name], [employee].[lastname], [schedule].[timeStart], [schedule].[timeEnd], [schedule].[data]" +
                    "FROM [employee], [schedule] " +
                    "WHERE [schedule].[employeeID]=[employee].[id] AND [employee].[id] = @id AND [schedule].[data] LIKE @date ", connection)
                {
                    CommandType = CommandType.Text
                };
                command.Parameters.AddWithValue("@id", id);
                command.Parameters.AddWithValue("@date", '%' + date.ToString("MM.yyyy"));
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    listBox.Add(new Employee(int.Parse(reader[0].ToString()), reader[1].ToString(), reader[2].ToString(), reader[3].ToString(), reader[4].ToString(), DateTime.Parse(reader[5].ToString())));
                }
            }
            return listBox;
        }
    }
}
