using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
namespace Assistent_Modul_Online
{
    public partial class EditZayav : Form
    {
        public EditZayav()
        {
            InitializeComponent();
        }
        string connstring = @"Server=.\SQLEXPRESS;Database=modul-online;Trusted_Connection=true;";
        
        public DataTable SqlCon(string sql)
        { //запрос с возвращенной таблицей
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            DataTable DT = new DataTable();
            SqlCommand cmd = new SqlCommand(sql, con);
            try
            {
                using (SqlDataReader dr = cmd.ExecuteReader())
            {

                DT.Load(dr);

            }
            con.Close();
            con.Dispose();
             }
            catch
             {
                 MessageBox.Show("Ошибка подключения к серверу");
             }
            return DT;
        }
        public void InsUpd(string sqlstring, string table, DataGridView drg)
        { //выполнение команды (добав. удал, измн)
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            try
            {
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = sqlstring;
                cmd.ExecuteNonQuery();
                con.Close();
                con.Dispose();
                drg.DataSource = SqlCon("select * from " + table+ " where MONTH(Дата) = '"+dateTimePicker3.Value.Month+"'");
                drg.Refresh();
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }

        private void GradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void EditZayav_Load(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewedtzav where MONTH(Дата)='" + dateTimePicker3.Value.Month + "'");

            SqlDataAdapter vid6 = new SqlDataAdapter("select * from Диллеры", connstring);
            DataTable dil2 = new DataTable();
            vid6.Fill(dil2);
            comboBox2.DataSource = dil2;
            comboBox2.DisplayMember = "Наименование";
            comboBox2.ValueMember = "id_dil";

            SqlDataAdapter vid4 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable dil = new DataTable();
            vid4.Fill(dil);
            comboBox1.DataSource = dil;
            comboBox1.DisplayMember = "ФИО";
            comboBox1.ValueMember = "id_sotr";

            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
                     
        }

        private void DataGridView2_Click(object sender, EventArgs e)
        {
            textBox1.Text = Convert.ToString(dataGridView2[2, dataGridView2.CurrentRow.Index].Value.ToString());//nomer
            textBox4.Text = Convert.ToString(dataGridView2[7, dataGridView2.CurrentRow.Index].Value.ToString());//stroki
            textBox5.Text = Convert.ToString(dataGridView2[8, dataGridView2.CurrentRow.Index].Value.ToString());//shtyki                      
            textBox2.Text = Convert.ToString(dataGridView2[4, dataGridView2.CurrentRow.Index].Value.ToString());//nachal
            textBox3.Text = Convert.ToString(dataGridView2[6, dataGridView2.CurrentRow.Index].Value.ToString());//vremya sbork
            comboBox2.Text = Convert.ToString(dataGridView2[1, dataGridView2.CurrentRow.Index].Value.ToString());//diler
            comboBox1.Text = Convert.ToString(dataGridView2[3, dataGridView2.CurrentRow.Index].Value.ToString());//sotrydn
            dateTimePicker1.Value = Convert.ToDateTime(dataGridView2[5, dataGridView2.CurrentRow.Index].Value.ToString());//data sbor
            checkBox1.Checked = Convert.ToBoolean(dataGridView2[9, dataGridView2.CurrentRow.Index].Value.ToString());//statys
            if (dataGridView2[10, dataGridView2.CurrentRow.Index].Value.ToString() == "")//otgryzka
            { }
            else
            { dateTimePicker2.Value = Convert.ToDateTime(dataGridView2[10, dataGridView2.CurrentRow.Index].Value.ToString()); };

        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        {
            int a = 0;

            if (checkBox1.Checked == true)
            {
                a = 1;
            }

            else { a = 0; };

            int p = Convert.ToInt32(dataGridView2[0, dataGridView2.CurrentRow.Index].Value);

            InsUpd("update Заявка set Статус="+a+", Номер='"+textBox1.Text+"', Строки='"+textBox4.Text+"', Штуки='"+textBox5.Text+"', Время_начала='"+textBox2.Text+"', Время_сборки='"+textBox3.Text+"', id_dil='"+comboBox2.SelectedValue+ "', id_sotr='" + comboBox1.SelectedValue+"', Дата_сборки='"+dateTimePicker1.Value+"', Отгрузка='"+dateTimePicker2.Value+"' where id_zaiav='"+p+"'", "viewedtzav" , dataGridView2);
            
        }

        private void DataGridView2_MultiSelectChanged(object sender, EventArgs e)
        {
            
            
        }

        private void DataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count >1)
            {
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                comboBox2.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
            else
            {
                textBox1.Enabled = true;
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                textBox5.Enabled = true;
                comboBox2.Enabled = true;              
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {

            if (dataGridView2.SelectedRows.Count > 0)
            {
                int ps;

                int p = Convert.ToInt32(dataGridView2[0, dataGridView2.CurrentRow.Index].Value);

                if (dataGridView2.SelectedRows.Count > 1)
                {
                    for (int i = 0; i < dataGridView2.SelectedRows.Count; i++)
                    {
                        ps = Convert.ToInt32(dataGridView2[0, dataGridView2.SelectedRows[i].Index].Value);
                        SqlCon("update Заявка set Статус='1' where id_zaiav=" + ps);
                    }
                    dataGridView2.DataSource = SqlCon("select * from viewedtzav where MONTH(Дата)='" + dateTimePicker3.Value.Month + "'");
                }
                else
                {
                    InsUpd("update Заявка set Статус='1' where id_zaiav=" + p, "viewedtzav", dataGridView2);
                }
            }
        }

        private void MaterialRaisedButton3_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView2[0, dataGridView2.CurrentRow.Index].Value);

            var result = MessageBox.Show("Удалить заявку: " + dataGridView2[2, dataGridView2.CurrentRow.Index].Value.ToString(), "Внимание",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                InsUpd("delete Заявка where id_zaiav=" + p, "viewedtzav", dataGridView2);
        }

        private void MaterialFlatButton1_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewedtzav where MONTH(Дата)='"+dateTimePicker3.Value.Month+"'");
        }

        private void MaterialFlatButton2_Click(object sender, EventArgs e)
        {
           dataGridView2.DataSource = SqlCon("select * from viewedtzav where Дата='" + dateTimePicker3.Value.ToShortDateString() + "'");
        }

        private void MaterialFlatButton3_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewedtzav where MONTH(Дата)='" + dateTimePicker3.Value.Month + "'");
        }

        private void DateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewedtzav where MONTH(Дата)='" + dateTimePicker3.Value.Month + "'");
        }

        private void TextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void TextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void TextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }
    }
}
