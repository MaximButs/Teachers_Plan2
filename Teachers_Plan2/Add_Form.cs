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

namespace Teachers_Plan2
{
    public partial class Add_Form : Form
    {

        DataBase dataBase = new DataBase();

        public Add_Form()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataBase.openConnection();

            var fio = textBox_fio2.Text;
            int time;
            var discip = textBox_discip2.Text;
            var semestr = textBox_semestr2.Text;
            var squad = textBox_squad2.Text;
            var test = textBox_test2.Text;

            if (int.TryParse(textBox_time2.Text, out time))
            {
                var addQuery = $"insert into Plan_db (fio_of, time_of, discipline, semestr, squad, test) values ('{fio}', '{time}', '{discip}', '{semestr}', '{squad}', '{test}')";

                var command = new SqlCommand(addQuery, dataBase.getConnection());
                command.ExecuteNonQuery();

                MessageBox.Show("Запись успешно создана!!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Количество часов должно иметь числовой формат!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            dataBase.closeConnection();

        }
    }
}
