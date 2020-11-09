using System.Data.SQLite;
using System.IO;

namespace SCUD
{
    public class Connection
    {
        public string Connect { get; } = @"Data Source=.\scud.db";
        public Connection()
        {
            if (!File.Exists(@"scud.db"))//поиск базы
            {
                SQLiteConnection.CreateFile(@"scud.db");//создание базы, если ее нет
                using (SQLiteConnection connection = new SQLiteConnection(Connect))
                {
                    SQLiteCommand command = new SQLiteCommand("CREATE TABLE [users] ( " +
                        "[id]    INTEGER NOT NULL UNIQUE, " +
                        "[login] TEXT NOT NULL," +
                        "[password]  TEXT NOT NULL," +
                        "[role]  NUMERIC NOT NULL," +
                        "[name]  TEXT," +
                        "[lastname]  TEXT," +
                        "PRIMARY KEY([id] AUTOINCREMENT));" +
                        "" +
                        "CREATE TABLE [position] ([id]    INTEGER NOT NULL UNIQUE," +
                        "[name]  TEXT," +
                        "PRIMARY KEY([id] AUTOINCREMENT));" +
                        "" +
                        "CREATE TABLE [departament] ([id] INTEGER NOT NULL UNIQUE," +
                        "[name]  TEXT NOT NULL," +
                        "PRIMARY KEY([id] AUTOINCREMENT));" +
                        "" +
                        "CREATE TABLE [workarea] (" +
                        "[id]    INTEGER NOT NULL UNIQUE," +
                        "[name]  TEXT NOT NULL," +
                        "PRIMARY KEY([id] AUTOINCREMENT));" +
                        "" +
                        "CREATE TABLE [employee] ([id]    INTEGER NOT NULL UNIQUE,    " +
                        "[name]  TEXT NOT NULL," +
                        "[lastname] TEXT NOT NULL," +
                        "[departamentID] INTEGER NOT NULL," +
                        "[positionID]    INTEGER NOT NULL," +
                        "[workareaID]    INTEGER NOT NULL," +
                        "PRIMARY KEY([id] AUTOINCREMENT)," +
                        "FOREIGN KEY([positionID]) REFERENCES position(id)," +
                        "FOREIGN KEY([departamentID]) REFERENCES departament(id)," +
                        "FOREIGN KEY([workareaID]) REFERENCES workarea(id));" +
                        "" +
                        "CREATE TABLE [schedule] ([id]    INTEGER NOT NULL UNIQUE," +
                        "[employeeID]    INTEGER NOT NULL," +
                        "[data]  TEXT NOT NULL," +
                        "[timeStart] TEXT NOT NULL," +
                        "[timeEnd]   TEXT," +
                        "PRIMARY KEY([id] AUTOINCREMENT)," +
                        "FOREIGN KEY([employeeID]) REFERENCES employee(id));" +
                        "" +
                        "INSERT INTO [users] ([login], [password], [role]) VALUES(@login, @password, @role);", connection);
                    command.Parameters.AddWithValue("@login", "admin");
                    command.Parameters.AddWithValue("@password", Encryption.GetHash("admin"));
                    command.Parameters.AddWithValue("@role", true);
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }


    }
}
