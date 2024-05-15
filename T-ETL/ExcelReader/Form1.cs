using System;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        private ComboBox cboSheet; // Declare the combo box variable

        public Form1()
        {
            InitializeComponent();
            InitializeControls();
            this.WindowState = FormWindowState.Maximized;
        }

        private void InitializeControls()
        {
            cboSheet = new ComboBox();
            cboSheet.Location = new System.Drawing.Point(10, 10); // Adjust the location as needed
            cboSheet.DropDownStyle = ComboBoxStyle.DropDownList; // Set dropdown style
            cboSheet.Width = 150; // Set width as needed
            this.Controls.Add(cboSheet); // Add combo box to form

            // Add items to the combo box
            cboSheet.Items.Add("Article");
            cboSheet.Items.Add("Achat");
            cboSheet.Items.Add("Vente");
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            try
            {
                // Get the selected sheet name
                string selectedSheet = cboSheet.SelectedItem.ToString();

                // Read data based on the selected sheet
                switch (selectedSheet)
                {
                    case "Article":
                        // Read articles from Excel
                        var articles = Helper.ReadArticles("registre.xls");
                        // Insert articles into SQL Server
                        Helper.InsertArticles(articles);
                        // Display articles in DataGridView
                        grd.DataSource = articles;
                        break;
                    case "Achat":
                        // Read achats from Excel
                        var achats = Helper.ReadAchats("registre.xls");
                        // Insert achats into SQL Server
                        Helper.InsertAchats(achats);
                        // Display achats in DataGridView
                        grd.DataSource = achats;
                        break;
                    case "Vente":
                        // Read ventes from Excel
                        var ventes = Helper.ReadVentes("registre.xls");
                        // Insert ventes into SQL Server
                        Helper.InsertVentes(ventes);
                        // Display ventes in DataGridView
                        grd.DataSource = ventes;
                        break;
                    default:
                        MessageBox.Show("Invalid sheet selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnBilan_Click(object sender, EventArgs e)
        {
            try
            {
                // Generate the bilan
                Helper.GenerateBilan();

                MessageBox.Show("Bilan generated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DataTable bilanData = Helper.GetBilanFromDatabase();

                // Display the bilan data in a DataGridView
                grd.DataSource = bilanData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while generating the bilan: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }

}
