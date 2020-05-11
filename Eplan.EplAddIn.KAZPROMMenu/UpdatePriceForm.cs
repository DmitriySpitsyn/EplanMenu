using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;

namespace Eplan.EplAddIn.KAZPROMMenu
{
    public partial class UpdatePriceForm : Form
    {
        public UpdatePriceForm()
        {
            InitializeComponent();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            label3.Enabled = true;
            label4.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            label3.Enabled = false;
            label4.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Enabled = true;
            progressBar1.Value = 10;
            string connectionString = "";
            if (radioButton1.Checked == true)
            {
                connectionString = @"Server=" + textBox1.Text + ";Database=ENGINEERING;Integrated Security=true;";

            }
            else if (radioButton2.Checked == true)
            {
                connectionString = @"Server=" + textBox1.Text + ";Database=ENGINEERING;User Id=" + textBox2.Text + ";Password=" + textBox3.Text + ";";
            }

            string sqlExpression = "update tblPart  SET tblPart.erpnr = tblPrices.code FROM tblPrices JOIN tblPart ON(tblPart.ordernr = tblPrices.article and tblPrices.article != '');";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    progressBar1.Value = 20;
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    progressBar1.Value += 20;
                    int number = command.ExecuteNonQuery();
                    progressBar1.Value += 20;
                }
                catch (SqlException ex)
                {
                    MessageBox.Show("Ошибка подлючения", "Ошибка");
                }


            }
            sqlExpression = "UPDATE tblPart SET salesprice1 = [tblPrices].[price1],[salesprice2] = [tblPrices].[price2] FROM[tblPrices] JOIN tblPart ON tblPart.erpnr = [tblPrices].[code]; ";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    progressBar1.Value += 10;
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    progressBar1.Value += 10;
                    int number = command.ExecuteNonQuery();
                    progressBar1.Value = 100;
                }
                catch (SqlException ex)
                {
                    MessageBox.Show("Ошибка подлючения", "Ошибка");
                }
            }

            progressBar1.Enabled = false;
        }
    }
}
