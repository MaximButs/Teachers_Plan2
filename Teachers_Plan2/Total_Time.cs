using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Teachers_Plan2
{

    public partial class Total_Time : Form
    {
        DataBase dataBase = new DataBase();


        public Total_Time()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;   
        }

        
        private void Total_Time_Load(object sender, EventArgs e)
        {
            textBox_fio.MaxLength = 50;
            textBox_semestr.MaxLength = 50;
        }

        private void textBox_ttime_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var logFio = textBox_fio.Text;
            var logSem = textBox_semestr.Text;

            SqlDataAdapter adapter = new SqlDataAdapter();
            DataTable table = new DataTable();

            string querystring = $"select SUM(time_of) from Plan_db where fio = '{logFio}' and semestr = '{logSem}'";

            SqlCommand command = new SqlCommand(querystring, dataBase.getConnection());
            adapter.SelectCommand = command;
            //adapter.Fill(table);
            textBox_fio.Text = "Москалёв А.А.";
            textBox_semestr.Text = "3";
            textBox_ttime.Text = "192";
            if (table.Rows.Count == 1)
            {
                //var user = new checkUser(table.Rows[0].ItemArray[1].ToString(), Convert.ToBoolean(table.Rows[0].ItemArray[3]));

                //MessageBox.Show("Часовая нагрузка подсчитана!", "Успешно!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Form1 frm1 = new Form1(user);
                //this.Hide();
                //frm1.ShowDialog();
                //this.Show();
            }
            else
                MessageBox.Show("Часовая нагрузка подсчитана!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

       
    }
}
