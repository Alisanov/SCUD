using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SQLite;
using System.Windows.Controls;

namespace SCUD
{
    class Directory
    {

        private Connection connect;

        private int id;
        private string content;

        public int ID { get { return id; } }
        public string Content { get { return content; } }
        public Directory()
        {
            this.id = -1;
            this.content = string.Empty;
            connect = new Connection();
        }
        public Directory(ref ListBox departamentList, ref ListBox workareaList, ref ListBox positionList)
        {
            this.id = 0;
            this.content = string.Empty;
            connect = new Connection();
            LoadDirectory("departament", ref departamentList);
            LoadDirectory("workarea", ref workareaList);
            LoadDirectory("position", ref positionList);
        }
        public Directory(int id, string content)
        {
            this.id = id;
            this.content = content;
        }
        public void SaveDirectory(string table, string text, ref ListBox listBox)//сохранение справочника
        {
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand("INSERT INTO [" + table + "] ([name]) VALUES(@name)", connection)
                {
                    CommandType = CommandType.Text
                };
                command.Parameters.AddWithValue("@name", text);
                command.ExecuteNonQuery();
            }

            LoadDirectory(table, ref listBox);
        }
        public void LoadDirectory(string table, ref ListBox listBox)//загрузка справочника
        {
            ObservableCollection<Directory> listDirectory = new ObservableCollection<Directory>();

            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand("SELECT * FROM [" + table + "] ORDER BY name", connection)
                {
                    CommandType = CommandType.Text
                };
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    listDirectory.Add(new Directory(int.Parse(reader["id"].ToString()), reader["name"].ToString()));
                }
            }
            listBox.ItemsSource = listDirectory;
        }
        public void DeleteDirectory(string table, ref ListBox listBox)//удлаение элемента справочника
        {
            Directory departament = (Directory)listBox.SelectedItem;
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand("DELETE FROM [" + table + "] WHERE [id]=@id", connection)
                {
                    CommandType = CommandType.Text
                };
                command.Parameters.AddWithValue("@id", departament.ID);
                command.ExecuteNonQuery();
            }
            LoadDirectory(table, ref listBox);
        }

        public void LoadComboBox(ref ComboBox comboBox, ref List<int> idComboBox, string table)
        {
            comboBox.Items.Clear();
            idComboBox.Clear();
            using (SQLiteConnection connection = new SQLiteConnection(connect.Connect))
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand("SELECT * FROM [" + table + "] ORDER BY name", connection);
                SQLiteDataReader read = command.ExecuteReader();
                while (read.Read())
                {
                    idComboBox.Add(Convert.ToInt32(read[0]));
                    comboBox.Items.Add(read[1]);
                }

                if (comboBox.Items.Count != 0)
                {
                    comboBox.Text = comboBox.Items[0].ToString();
                }
                else
                {
                    comboBox.Text = "В справочнике отуствуют данные";
                }
            }
        }
    }
}
