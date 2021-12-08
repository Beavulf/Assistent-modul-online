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
    public partial class StatsZayav : Form
    {
        public StatsZayav()
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
            //try
            //{
                using (SqlDataReader dr = cmd.ExecuteReader())
                {

                    DT.Load(dr);

                }
                con.Close();
                con.Dispose();
        //}
        //    catch
        //    {
        //        MessageBox.Show("Ошибка подключения к серверу");
        //    }
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
        private void StatsZayav_Load(object sender, EventArgs e)
        {
            SqlDataAdapter vid6 = new SqlDataAdapter("select * from Диллеры", connstring);
            DataTable dil2 = new DataTable();
            vid6.Fill(dil2);
            comboBox2.DataSource = dil2;
            comboBox2.DisplayMember = "Наименование";
            comboBox2.ValueMember = "id_dil";


            dataGridView3.DataSource = SqlCon("select Диллеры.Наименование as Дилер, COUNT(id_zaiav) as Колво, concat(round((COUNT(id_zaiav)/convert(float,(select count(*) from Заявка))*100),2),'%') as проценты from Диллеры join Заявка on Заявка.id_dil = Диллеры.id_dil group by Диллеры.Наименование");

            dataGridView2.DataSource = SqlCon("select * from viewzav where MONTH(Дата)='"+dateTimePicker4.Value.Month+"'");

            SqlDataAdapter vid4 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable dil = new DataTable();
            vid4.Fill(dil);
            comboBox4.DataSource = dil;
            comboBox4.DisplayMember = "ФИО";
            comboBox4.ValueMember = "id_sotr";
        }

        private void MaterialFlatButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MaterialFlatButton1_Click_1(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewzav where MONTH(Дата)='" + dateTimePicker4.Value.Month + "'");
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {          
            switch (comboBox2.Text.ToString())
            {
                case "Все":
                    dataGridView2.DataSource = SqlCon("select * from viewzav");
                    
                    break;
                default:
                    dataGridView2.DataSource = SqlCon("select * from viewzav where Дилер='" + comboBox2.Text + "'");
                    
                    break;
            }
            
        }

        private void ComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.CurrentCell = null;
                dataGridView2.Rows[i].Visible = false;
                for (int c = 0; c < dataGridView2.Columns.Count; c++)
                {
                    if (dataGridView2[c, i].Value.ToString() == comboBox4.Text)
                    {
                        dataGridView2.Rows[i].Visible = true;
                        break;
                    }
                }
            }
        }

        private void MaterialSingleLineTextField1_Click(object sender, EventArgs e)
        {

        }

        private void MaterialSingleLineTextField1_TextChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewzav where Номер = '"+materialSingleLineTextField1.Text+"'");
        }

        private void DateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            

            if (materialRadioButton1.Checked == true)
            {
                dataGridView2.DataSource = SqlCon("select * from viewzav where Дата='"+dateTimePicker4.Value.ToShortDateString()+"'");
            }
            else { dataGridView2.DataSource = SqlCon("select * from viewzav where MONTH(Дата)='" + dateTimePicker4.Value.Month + "'"); }
        }

        private void MaterialLabel2_Click(object sender, EventArgs e)
        {

        }

        private void MaterialLabel3_Click(object sender, EventArgs e)
        {

        }

        private void GradientPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void MaterialFlatButton3_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            if (materialRadioButton1.Checked)
            {
                SqlCommand cmd = new SqlCommand("CalcZayavStatsFull", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@data", SqlDbType.Date).Value = dateTimePicker4.Value;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@atime", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@psborki", SqlDbType.Float));
                cmd.Parameters.Add(new SqlParameter("@shtuki", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@zcount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@scount", SqlDbType.Int));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@atime"].Direction = ParameterDirection.Output;
                cmd.Parameters["@psborki"].Direction = ParameterDirection.Output;
                cmd.Parameters["@shtuki"].Direction = ParameterDirection.Output;
                cmd.Parameters["@zcount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@scount"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label24.Text = cmd.Parameters["@count"].Value.ToString();
                label25.Text = cmd.Parameters["@acount"].Value.ToString() + " строк(и)";
                label26.Text = cmd.Parameters["@atime"].Value.ToString() + " минут(ы)";
                label27.Text = cmd.Parameters["@psborki"].Value.ToString() + " %";
                label28.Text = cmd.Parameters["@shtuki"].Value.ToString() + " штук(и)";
                label29.Text = cmd.Parameters["@zcount"].Value.ToString();
                label30.Text = cmd.Parameters["@scount"].Value.ToString();
                string report1 = "select * from viewzav where Дата='" + dateTimePicker4.Value + "'";
                dataGridView2.DataSource = SqlCon(report1);
                con.Close();
                con.Dispose();
            }
            else
            { MessageBox.Show("Выберите дату", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }          }

        private void MaterialFlatButton4_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            string str = dateTimePicker4.Value.Month.ToString();
            if (materialRadioButton1.Checked)
            {
                SqlCommand cmd = new SqlCommand("CalcZayavStatsFullMonth", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@data", SqlDbType.VarChar,3).Value = str;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@atime", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@psborki", SqlDbType.Float));
                cmd.Parameters.Add(new SqlParameter("@shtuki", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@zcount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@scount", SqlDbType.Int));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@atime"].Direction = ParameterDirection.Output;
                cmd.Parameters["@psborki"].Direction = ParameterDirection.Output;
                cmd.Parameters["@shtuki"].Direction = ParameterDirection.Output;
                cmd.Parameters["@zcount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@scount"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label24.Text = cmd.Parameters["@count"].Value.ToString();
                label25.Text = cmd.Parameters["@acount"].Value.ToString() + " строк(и)";
                label26.Text = cmd.Parameters["@atime"].Value.ToString() + " минут(ы)";
                label27.Text = cmd.Parameters["@psborki"].Value.ToString() + " %";
                label28.Text = cmd.Parameters["@shtuki"].Value.ToString() + " штук(и)";
                label29.Text = cmd.Parameters["@zcount"].Value.ToString();
                label30.Text = cmd.Parameters["@scount"].Value.ToString();
                string report = "select * from viewzav where MONTH(Дата)='"+str+"'";
                dataGridView2.DataSource = SqlCon(report);
                con.Close();
                con.Dispose();
            }
            else
            { MessageBox.Show("Выберите дату", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }
        }

        private void DateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            if (materialRadioButton2.Checked == true)
            {
                dataGridView2.DataSource = SqlCon("select * from viewzav where Отгрузка = '" + dateTimePicker5.Value.ToShortDateString() + "'");
            }
        }
    }
}

