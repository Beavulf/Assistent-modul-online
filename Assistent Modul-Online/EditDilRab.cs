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
    public partial class EditDilRab : Form
    {
        public EditDilRab()
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
                drg.DataSource = SqlCon("select * from " + table);
                drg.Refresh();
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void EditDilRab_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = SqlCon("select * from Другие_работы");

            dataGridView2.DataSource = SqlCon("select * from Диллеры");

            dataGridView3.DataSource = SqlCon("select * from viewvrabote where MONTH(Дата)='"+dateTimePicker1.Value.Month+"'");

            SqlDataAdapter vid4 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable dil = new DataTable();
            vid4.Fill(dil);
            comboBox2.DataSource = dil;
            comboBox2.DisplayMember = "ФИО";
            comboBox2.ValueMember = "id_sotr";

            SqlDataAdapter vid2 = new SqlDataAdapter("select * from Другие_работы", connstring);
            DataTable rab = new DataTable();
            vid2.Fill(rab);
            comboBox1.DataSource = rab;
            comboBox1.DisplayMember = "Название";
            comboBox1.ValueMember = "id_rab";
        }

        private void DataGridView3_Click(object sender, EventArgs e)
        {
            comboBox1.Text = Convert.ToString(dataGridView3[1, dataGridView3.CurrentRow.Index].Value.ToString());
            comboBox2.Text = Convert.ToString(dataGridView3[4, dataGridView3.CurrentRow.Index].Value.ToString());
        }

        private void MaterialRaisedButton7_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView3[0, dataGridView3.CurrentRow.Index].Value);

            InsUpd("update Вработе set id_rab='"+comboBox1.SelectedValue+"' , id_sotr='"+comboBox2.SelectedValue+"' where id_vrab='" + p + "'", "viewvrabote where MONTH(Дата)='" + dateTimePicker1.Value.Month + "'", dataGridView3);
        }

        private void MaterialRaisedButton8_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView3[0, dataGridView3.CurrentRow.Index].Value);

            var result = MessageBox.Show("Удалить запись о работе №: " + dataGridView3[0, dataGridView3.CurrentRow.Index].Value.ToString(), "Внимание",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                InsUpd("delete Вработе where id_vrab=" + p, "viewvrabote where MONTH(Дата)='" + dateTimePicker1.Value.Month + "'", dataGridView3);
        }

        private void DataGridView1_Click(object sender, EventArgs e)
        {
            materialSingleLineTextField1.Text = dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString();
        }

        private void MaterialRaisedButton3_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentRow.Index].Value);

            var result = MessageBox.Show("Удалить работу: " + dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString(), "Внимание",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                InsUpd("delete Другие_работы where id_rab=" + p, "Другие_работы", dataGridView1);
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {
            InsUpd("insert into Другие_работы (Название) values ('"+materialSingleLineTextField1.Text+"')", "Другие_работы", dataGridView1);
        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentRow.Index].Value);

            InsUpd("update Другие_работы set Название='" + materialSingleLineTextField1.Text + "' where id_rab='"+p+"'", "Другие_работы", dataGridView1);
        }

        private void DataGridView2_Click(object sender, EventArgs e)
        {
            materialSingleLineTextField2.Text = dataGridView2[1, dataGridView2.CurrentRow.Index].Value.ToString();
            materialSingleLineTextField3.Text = dataGridView2[2, dataGridView2.CurrentRow.Index].Value.ToString();
        }

        private void MaterialRaisedButton4_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView2[0, dataGridView2.CurrentRow.Index].Value);

            var result = MessageBox.Show("Удалить дилера: " + dataGridView2[1, dataGridView2.CurrentRow.Index].Value.ToString(), "Внимание",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                InsUpd("delete Диллеры where id_dil=" + p, "Диллеры", dataGridView2);
        }

        private void MaterialRaisedButton6_Click(object sender, EventArgs e)
        {
            InsUpd("insert into Диллеры (Наименование, Расшифровка) values ('" + materialSingleLineTextField2.Text + "', '" + materialSingleLineTextField3.Text + "')", "Диллеры", dataGridView2);
        }

        private void MaterialRaisedButton5_Click(object sender, EventArgs e)
        {
            int p = Convert.ToInt32(dataGridView2[0, dataGridView2.CurrentRow.Index].Value);

            InsUpd("update Диллеры set Наименование='" + materialSingleLineTextField2.Text + "', Расшифровка='" + materialSingleLineTextField3.Text + "' where id_dil='" + p + "'", "Диллеры", dataGridView2);
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView3.DataSource = SqlCon("select * from viewvrabote where MONTH(Дата)='" + dateTimePicker1.Value.Month + "'");
        }

        private void MaterialSingleLineTextField2_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void MaterialSingleLineTextField3_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }
    }
}
