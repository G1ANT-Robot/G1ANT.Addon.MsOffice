using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Forms
{
    public partial class DataTableForm : Form
    {
        public DataTableForm()
        {
            InitializeComponent();
        }

        public void LoadData(OleDbConnection dbConnection, string tableName)
        {
            var ds = new DataSet();
            using (var command = new OleDbCommand($"select * from {tableName}", dbConnection))
            {
                var adp = new OleDbDataAdapter(command);
                adp.Fill(ds, tableName);
                dataGridView1.DataSource = ds.Tables[tableName];
            }
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
