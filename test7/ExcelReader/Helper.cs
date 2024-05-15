using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;
namespace ExcelReader
{
    public class Helper
    {
        private const string connectionString = "Data Source=LEETBOOK\\SQLEXPRESS;Initial Catalog=ETL;Integrated Security=True"; // Adjust connection string for SQL Server

        public static List<Article> ReadArticles(string relativePath)
        {
            return ReadDataFromSheet<Article>(relativePath, 1, CreateArticle);
        }

        public static List<Achat> ReadAchats(string relativePath)
        {
            return ReadDataFromSheet<Achat>(relativePath, 2, CreateAchat);
        }

        public static List<Vente> ReadVentes(string relativePath)
        {
            return ReadDataFromSheet<Vente>(relativePath, 3, CreateVente);
        }

        private static List<T> ReadDataFromSheet<T>(string relativePath, int sheetIndex, Func<Excel.Range, T> createObject)
        {
            List<T> result = new List<T>();
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(basePath, relativePath);
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[sheetIndex]; // Select the specified sheet
            Excel.Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            for (int ligne = 2; ligne <= rowCount; ligne++)
            {
                bool hasNull = false;

                // Check if any cell in the row is null
                for (int col = 1; col <= colCount; col++)
                {
                    if (range.Cells[ligne, col].Value2 == null)
                    {
                        hasNull = true;
                        break;
                    }
                }

                // If any cell is null, skip creating an object for this row
                if (hasNull)
                {
                    continue;
                }

                T obj = createObject(range.Rows[ligne]);
                result.Add(obj);
            }

            workbook.Close();
            excelApp.Quit();

            return result;
        }

        private static Article CreateArticle(Excel.Range row)
        {
            return new Article(
                row.Cells[1].Value2.ToString(),
                row.Cells[2].Value2.ToString(),
                row.Cells[3].Value2.ToString());
        }

        private static Achat CreateAchat(Excel.Range row)
        {
            double serialDate = (double)row.Cells[4].Value2;
            DateTime dateValue = DateTime.FromOADate(serialDate);

            return new Achat(
                row.Cells[1].Value2.ToString(),
                row.Cells[2].Value2.ToString(),
                row.Cells[3].Value2.ToString(),
                dateValue);
        }

        private static Vente CreateVente(Excel.Range row)
        {
            double serialDate = (double)row.Cells[4].Value2;
            DateTime dateValue = DateTime.FromOADate(serialDate);

            return new Vente(
                row.Cells[1].Value2.ToString(),
                row.Cells[2].Value2.ToString(),
                row.Cells[3].Value2.ToString(),
                dateValue);
        }


        // LOAD

        public static void InsertArticles(List<Article> articles)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Check if the table exists, if not, create it
                string createTableQuery = @"
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'articles')
            BEGIN
                CREATE TABLE articles (
                    id INT PRIMARY KEY,
                    libelle NVARCHAR(MAX),
                    pu INT
                )
            END";

                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                {
                    createTableCommand.ExecuteNonQuery();
                }

                // Now insert the data
                foreach (Article article in articles)
                {
                    string checkQuery = "SELECT COUNT(*) FROM articles WHERE id = @id";
                    using (SqlCommand checkCommand = new SqlCommand(checkQuery, connection))
                    {
                        checkCommand.Parameters.AddWithValue("@id", article.Id);
                        int count = (int)checkCommand.ExecuteScalar();
                        if (count == 0)
                        {
                            // If the article doesn't exist, insert it
                            string insertQuery = "INSERT INTO articles (id, libelle, pu) VALUES (@id, @libelle, @pu)";
                            using (SqlCommand insertCommand = new SqlCommand(insertQuery, connection))
                            {
                                insertCommand.Parameters.AddWithValue("@id", article.Id);
                                insertCommand.Parameters.AddWithValue("@libelle", article.Libelle);
                                insertCommand.Parameters.AddWithValue("@pu", article.PU);
                                insertCommand.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }

        public static void InsertAchats(List<Achat> achats)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = @"
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'achats')
            BEGIN
                CREATE TABLE achats (
                    num INT PRIMARY KEY,
                    art_id INT,
                    qte INT,
                    date DATETIME
                )
            END";
                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                {
                    createTableCommand.ExecuteNonQuery();
                }
                foreach (Achat achat in achats)
                {
                    string checkQuery = "SELECT COUNT(*) FROM achats WHERE num = @num";
                    using (SqlCommand checkCommand = new SqlCommand(checkQuery, connection))
                    {
                        checkCommand.Parameters.AddWithValue("@num", achat.Num);
                        int count = (int)checkCommand.ExecuteScalar();
                        if (count == 0)
                        {
                            string query = "INSERT INTO achats (num, art_id, qte, date) VALUES (@num, @art_id, @qte, @date)";
                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@num", achat.Num);
                                command.Parameters.AddWithValue("@art_id", achat.Art_Id);
                                command.Parameters.AddWithValue("@qte", achat.Qte);
                                command.Parameters.AddWithValue("@date", achat.Date);
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }

        public static void InsertVentes(List<Vente> ventes)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = @"
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'ventes')
            BEGIN
                CREATE TABLE ventes (
                    num INT PRIMARY KEY,
                    art_id INT,
                    qte INT,
                    date DATETIME
                )
            END";
                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                {
                    createTableCommand.ExecuteNonQuery();
                }
                foreach (Vente vente in ventes)
                {
                    string checkQuery = "SELECT COUNT(*) FROM ventes WHERE num = @num";
                    using (SqlCommand checkCommand = new SqlCommand(checkQuery, connection))
                    {
                        checkCommand.Parameters.AddWithValue("@num", vente.Num);
                        int count = (int)checkCommand.ExecuteScalar();
                        if (count == 0)
                        {
                            string query = "INSERT INTO ventes (num, art_id, qte, date) VALUES (@num, @art_id, @qte, @date)";
                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@num", vente.Num);
                                command.Parameters.AddWithValue("@art_id", vente.Art_Id);
                                command.Parameters.AddWithValue("@qte", vente.Qte);
                                command.Parameters.AddWithValue("@date", vente.Date);
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }


        // BILAN

        public static void GenerateBilan()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = @"
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'bilan')
            BEGIN
                CREATE TABLE bilan (
                    id INT PRIMARY KEY IDENTITY,
                    art_id INT,
                    qte_actuelle INT
                )
            END";
                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                {
                    createTableCommand.ExecuteNonQuery();
                }
            }
            ClearBilanTable();

            List<Achat> achats = GetAchatsFromDatabase();
            List<Vente> ventes = GetVentesFromDatabase();

            Dictionary<int, int> totalQuantity = CalculateTotalQuantity(achats, ventes);

            InsertBilanEntries(totalQuantity);
        }

        private static void ClearBilanTable()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string clearTableQuery = "DELETE FROM bilan";
                using (SqlCommand command = new SqlCommand(clearTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private static List<Achat> GetAchatsFromDatabase()
        {
            List<Achat> achats = new List<Achat>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string selectQuery = "SELECT * FROM achats";
                using (SqlCommand command = new SqlCommand(selectQuery, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Achat achat = new Achat
                            {
                                Num = Convert.ToInt32(reader["num"]),
                                Art_Id = Convert.ToInt32(reader["art_id"]),
                                Qte = Convert.ToInt32(reader["qte"]),
                                Date = Convert.ToDateTime(reader["date"])
                            };
                            achats.Add(achat);
                        }
                    }
                }
            }
            return achats;
        }

        private static List<Vente> GetVentesFromDatabase()
        {
            List<Vente> ventes = new List<Vente>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string selectQuery = "SELECT * FROM ventes";
                using (SqlCommand command = new SqlCommand(selectQuery, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Vente vente = new Vente
                            {
                                Num = Convert.ToInt32(reader["num"]),
                                Art_Id = Convert.ToInt32(reader["art_id"]),
                                Qte = Convert.ToInt32(reader["qte"]),
                                Date = Convert.ToDateTime(reader["date"])
                            };
                            ventes.Add(vente);
                        }
                    }
                }
            }
            return ventes;
        }

        private static Dictionary<int, int> CalculateTotalQuantity(List<Achat> achats, List<Vente> ventes)
        {
            Dictionary<int, int> totalQuantity = new Dictionary<int, int>();

            // Calculate total quantity for each article
            foreach (var achat in achats)
            {
                if (!totalQuantity.ContainsKey(achat.Art_Id))
                {
                    totalQuantity[achat.Art_Id] = 0;
                }
                totalQuantity[achat.Art_Id] += achat.Qte;
            }

            foreach (var vente in ventes)
            {
                if (!totalQuantity.ContainsKey(vente.Art_Id))
                {
                    totalQuantity[vente.Art_Id] = 0;
                }
                totalQuantity[vente.Art_Id] -= vente.Qte;
            }

            return totalQuantity;
        }

        private static void InsertBilanEntries(Dictionary<int, int> totalQuantity)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                foreach (var kvp in totalQuantity)
                {
                    string insertQuery = "INSERT INTO bilan (art_id, qte_actuelle) VALUES (@art_id, @qte_actuelle)";
                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@art_id", kvp.Key);
                        command.Parameters.AddWithValue("@qte_actuelle", kvp.Value);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public static DataTable GetBilanFromDatabase()
        {
            DataTable bilanData = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT * FROM bilan";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(bilanData);
                    }
                }
            }

            return bilanData;
        }

    }
}
    