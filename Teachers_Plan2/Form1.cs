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
    enum RowState
    {
        Existed,
        New,
        Modified,
        ModifiedNew,
        Deleted
    }
    
    public partial class Form1 : Form
    {
        private readonly checkUser _user;
        
        DataBase dataBase = new DataBase();

        int selectedRow;
        public Form1(checkUser user)
        {
            _user = user;
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;   
        }

        

        private void IsAdmin()
        {
            управлениеToolStripMenuItem.Enabled = _user.IsAdmin;
            btnDelete.Enabled = _user.IsAdmin;
            btnChange.Enabled = _user.IsAdmin;
        }

        private void CreateColumns()
        {
            dataGridView1.Columns.Add("id", "id");
            dataGridView1.Columns.Add("fio_of", "Преподаватель");
            dataGridView1.Columns.Add("time_of", "Кол-во часов");
            dataGridView1.Columns.Add("discipline", "Дисциплина");
            dataGridView1.Columns.Add("semestr", "Семестр");
            dataGridView1.Columns.Add("squad", "Группа");
            dataGridView1.Columns.Add("test", "Форма контроля");
            dataGridView1.Columns.Add("IsNew", String.Empty);
            dataGridView1.Columns["IsNew"].Visible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void ClearFields()
        {
            textBox_id.Text = "";
            textBox_fio.Text = "";
            textBox_time.Text = "";
            textBox_discip.Text = "";
            textBox_semestr.Text = "";
            textBox_squad.Text = "";
            textBox_test.Text = "";
        }

        private void ReadSingleRow(DataGridView dgw, IDataRecord record)
        {
            dgw.Rows.Add(record.GetInt32(0), record.GetString(1), record.GetInt32(2), record.GetString(3), record.GetInt32(4), record.GetString(5), record.GetString(6), RowState.ModifiedNew);
        }

        private void RefreshDataGrid(DataGridView dgw)
        {
            dgw.Rows.Clear();

            string queryString = $"select * from Plan_db";

            SqlCommand command = new SqlCommand(queryString, dataBase.getConnection());

            dataBase.openConnection();

            SqlDataReader reader = command.ExecuteReader();

            while(reader.Read())
            {
                ReadSingleRow(dgw, reader);
            }
            reader.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtUserStatus.Text = $"{_user.Login}: {_user.Status}";
            IsAdmin();
            CreateColumns();
            RefreshDataGrid(dataGridView1);
            
        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;

            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[selectedRow];

                textBox_id.Text = row.Cells[0].Value.ToString();
                textBox_fio.Text = row.Cells[1].Value.ToString();
                textBox_time.Text = row.Cells[2].Value.ToString();
                textBox_discip.Text = row.Cells[3].Value.ToString();
                textBox_semestr.Text = row.Cells[4].Value.ToString();
                textBox_squad.Text = row.Cells[5].Value.ToString();
                textBox_test.Text = row.Cells[6].Value.ToString();
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            ClearFields();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshDataGrid(dataGridView1);
            ClearFields();

        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Add_Form addfrm = new Add_Form();
            addfrm.Show();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            deleteRow();
            ClearFields();
        }

        private void Change()
        {
            var selectedRowIndex = dataGridView1.CurrentCell.RowIndex;

            var id = textBox_id.Text;
            var fio = textBox_fio.Text;
            int time;
            var discip = textBox_discip.Text;
            var semestr = textBox_semestr.Text;
            var squad = textBox_squad.Text;
            var test = textBox_test.Text;

            if (dataGridView1.Rows[selectedRowIndex].Cells[0].Value.ToString() != string.Empty)
            {
                if(int.TryParse(textBox_time.Text, out time))
                {
                    dataGridView1.Rows[selectedRowIndex].SetValues(id, fio, time, discip, semestr, squad, test);
                    dataGridView1.Rows[selectedRowIndex].Cells[7].Value = RowState.Modified;
                }
                else
                {
                    MessageBox.Show("Количество часов должно иметь числовой формат!");
                }
            }
        }
        
        private void btnChange_Click(object sender, EventArgs e)
        {
            Change();
            ClearFields();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            updateRow();
        }

        private void Search(DataGridView dgw)
        {
            dgw.Rows.Clear();

            string searchString = $"select * from Plan_db where concat (id, fio_of, time_of, discipline, semestr, squad, test) like '%" + textBox_search.Text + "%'";

            SqlCommand com = new SqlCommand(searchString, dataBase.getConnection());

            dataBase.openConnection();

            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                ReadSingleRow(dgw, read);
            }

            read.Close();
        }

        private void deleteRow()
        {
            int index = dataGridView1.CurrentCell.RowIndex;

            dataGridView1.Rows[index].Visible = false;

            if (dataGridView1.Rows[index].Cells[0].Value.ToString() == string.Empty)
            {
                dataGridView1.Rows[index].Cells[7].Value = RowState.Deleted;
                return;
            }
            dataGridView1.Rows[index].Cells[7].Value = RowState.Deleted;
        }

        private void updateRow()
        {
            dataBase.openConnection();

            for(int index = 0; index < dataGridView1.Rows.Count; index++)
            {
                var rowState = (RowState)dataGridView1.Rows[index].Cells[7].Value;

                if (rowState == RowState.Existed)
                    continue;

                if(rowState == RowState.Deleted)
                {
                    var id = Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value);
                    var deleteQuery = $"delete from Plan_db where id = {id}";

                    var command = new SqlCommand(deleteQuery, dataBase.getConnection());
                    command.ExecuteNonQuery();
                }
                if(rowState == RowState.Modified)
                {
                    var id = dataGridView1.Rows[index].Cells[0].Value.ToString();
                    var fio = dataGridView1.Rows[index].Cells[1].Value.ToString();
                    var time = dataGridView1.Rows[index].Cells[2].Value.ToString();
                    var discip = dataGridView1.Rows[index].Cells[3].Value.ToString();
                    var semestr = dataGridView1.Rows[index].Cells[4].Value.ToString();
                    var squad = dataGridView1.Rows[index].Cells[5].Value.ToString();
                    var test = dataGridView1.Rows[index].Cells[6].Value.ToString();

                    var changeQuery = $"update Plan_db set fio_of = '{fio}', time_of = '{time}', discipline = '{discip}', semestr = '{semestr}', squad = '{squad}', test = '{test}' where id = '{id}'";

                    var command = new SqlCommand(changeQuery, dataBase.getConnection());
                    command.ExecuteNonQuery();
                }
            }
            dataBase.closeConnection();
        }
        
        private void textBox_search_TextChanged(object sender, EventArgs e)
        {
            Search(dataGridView1);
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();

            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;

            int i, j;
            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 2; j++)
                {
                    wsh.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                    wsh.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                }
            }
            exApp.Visible = true;
            exApp.Columns.AutoFit();
        }

        private void управлениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Panel_Admin frm_Panel = new Panel_Admin();
            frm_Panel.Show();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Программное средство формирования индивидуального плана преподавателя." +
                "\n\nCopyright © Буценко М.Н. \n\nAll Rights Reserved 2022. Все права защищены.", "О программе", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Текстовый документ (*.txt)|*.txt|Все файлы (*.*)|*.*";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter sw = new StreamWriter(saveFileDialog.FileName, false, Encoding.Unicode);
                    try
                    {
                        List<int> col_n = new List<int>();
                        foreach (DataGridViewColumn col in dataGridView1.Columns)
                            if (col.Visible)
                            {
                                sw.Write(col.HeaderText + "\t");
                                col_n.Add(col.Index);
                            }
                        sw.WriteLine();
                        int x = dataGridView1.RowCount;
                        if (dataGridView1.AllowUserToAddRows) x--;

                        for (int i = 0; i < x; i++)
                        {
                            for (int y = 0; y < col_n.Count; y++)
                                sw.Write(dataGridView1[col_n[y], i].Value + "\t");
                            sw.Write(" \r\n");
                        }
                        sw.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    }

                }
            }
        }

        private void итогToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Total_Time frm_Total = new Total_Time();
            frm_Total.Show();
        }

        private void оКафедреToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Кафедра микропроцессорных систем и сетей была основана в 1995 году и  обеспечивает проведение занятий " +
                "со слушателями переподготовки вечерней и заочной форм получения образования по специальностям: " +
                "«Программное обеспечение информационных систем» (вечерняя и заочная форма получения образования); " +
                "«Web-дизайн и компьютерная графика» (вечерняя и заочная форма получения образования); " +
                "«Тестирование программного обеспечения» (заочная форма получения образования), " +
                "а также учавствует в повышении квалификации и проведении обучающих курсов в области информационных технологий.\n\n" +
                "Кроме учебной работы на кафедре проводятся исследования по госбюджетной теме «Дополнительное образование взрослых " +
                "в сфере IT и цифровизация в рамках концепции Университет 3.0». " +
                "Преподаватели кафедры принимают участие в международных и республиканских научных конференциях.\n\n" +
                "На кафедре работают 4 кандидата наук (доцента), 6 старших преподавателей, 2 ассистента." +
                "Преподаватели кафедры являются сертифицированными специалистами в области компьютерных сетей (CISCO), " +
                "виртуализации (VMware) и управления проектами (SCRUM, Agile)\n\n" +
                "Кафедра находится в 7 корпусе БГУИР (аудитория №606). \n\n" +
                "Адрес в сети Интернет: https://iti.bsuir.by/department/3", "Кафедра микропроцессорных систем и сетей", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
