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
using MaterialSkin;
using MaterialSkin.Controls;

namespace Assistent_Modul_Online
{
    public partial class Report : MaterialForm
    {
        public Report()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue800, Primary.Blue400, Accent.LightBlue200, TextShade.WHITE);
        }
        //строка подключения
        string connstring = @"Server=.\SQLEXPRESS;Database=modul-online;Trusted_Connection=true;";

        public void InsUpd(string sqlstring, string table, DataGridView drg)
        { //выполнение команды (добав. удал, измн) (строка запроса, таблица, куда отображать)
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            try
            {
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = sqlstring;
            int a = cmd.ExecuteNonQuery();
            con.Close();
            drg.DataSource = SqlCon("select * from " + table);
            drg.Refresh();
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }

        public DataTable SqlCon(string sql)
        { //запрос с возвращенной таблицей
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            DataTable DT = new DataTable();
            SqlDataAdapter adap = new SqlDataAdapter(sql, con);
            SqlCommand cmd = new SqlCommand(sql, con);
            try
            {
                using (SqlDataReader dr = cmd.ExecuteReader())
                {

                    DT.Load(dr);

                }
                con.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к серверу");
            }
            return DT;
        }
        

        private void Report_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            SqlDataAdapter vid1 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable fam3 = new DataTable();
            vid1.Fill(fam3);
            comboBox1.DataSource = fam3;
            comboBox1.DisplayMember = "ФИО";
            comboBox1.ValueMember = "id_sotr";
        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        { //расчет строк

            
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {

        }

        private void MaterialRaisedButton1_Click_1(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            if (materialCheckBox1.Checked == true)
            {
                string str = dateTimePicker1.Value.Month.ToString();
                SqlCommand cmd = new SqlCommand("CalcReportMonth", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@sotr", SqlDbType.Int).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@data", SqlDbType.VarChar, 3).Value = str;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@mcount", SqlDbType.Int));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@mcount"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label6.Text = cmd.Parameters["@count"].Value.ToString();
                label7.Text = cmd.Parameters["@acount"].Value.ToString();
                label8.Text = cmd.Parameters["@mcount"].Value.ToString();
                string report = "select * from viewreport where ФИО='" + comboBox1.Text + "' and MONTH(Дата) =" + str;
                dataGridView1.DataSource = SqlCon(report);
            }
            else
            {
                SqlCommand cmd = new SqlCommand("CalcReport", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@sotr", SqlDbType.Int).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@data", SqlDbType.Date).Value = dateTimePicker1.Value;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@mcount", SqlDbType.Int));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@mcount"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label6.Text = cmd.Parameters["@count"].Value.ToString();
                label7.Text = cmd.Parameters["@acount"].Value.ToString();
                label8.Text = cmd.Parameters["@mcount"].Value.ToString();
                string report1 = "select * from viewreport where Дата='" + dateTimePicker1.Value + "' and ФИО='" + comboBox1.Text + "'";
                dataGridView1.DataSource = SqlCon(report1);
            }
            con.Close();
        }
    }
}
