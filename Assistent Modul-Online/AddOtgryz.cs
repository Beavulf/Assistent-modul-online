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
    public partial class AddOtgryz : MaterialForm
    {
        public AddOtgryz()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue800, Primary.Blue400, Accent.LightBlue200, TextShade.WHITE);
        }
        string connstring = @"Server=.\SQLEXPRESS;Database=modul-online;Trusted_Connection=true;";
        int p;
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
                drg.DataSource = SqlCon(table);
                drg.Refresh();
        }
            catch
            {
                MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
}

        private void AddOtgryz_Load(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewedtzav where Отгрузка is null");

            SqlDataAdapter vid6 = new SqlDataAdapter("select * from Диллеры", connstring);
            DataTable dil2 = new DataTable();
            vid6.Fill(dil2);
            comboBox1.DataSource = dil2;
            comboBox1.DisplayMember = "Наименование";
            comboBox1.ValueMember = "id_dil";
        }

        private void MaterialFlatButton1_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                int ps;
                p = Convert.ToInt32(dataGridView2[0, dataGridView2.CurrentRow.Index].Value);
                if (dataGridView2.SelectedRows.Count > 1)
                {
                    for (int i = 0; i < dataGridView2.SelectedRows.Count; i++)
                    {
                        ps = Convert.ToInt32(dataGridView2[0, dataGridView2.SelectedRows[i].Index].Value);
                        SqlCon("update Заявка set Отгрузка=FORMAT(getdate(),'yyyy-MM-dd') where id_zaiav=" + ps);
                    }
                    dataGridView2.DataSource = SqlCon("select * from viewedtzav where Отгрузка is null");
                }
                else
                {
                    InsUpd("update Заявка set Отгрузка=FORMAT(getdate(),'yyyy-MM-dd') where id_zaiav=" + p, "select * from viewedtzav where Отгрузка is null", dataGridView2);
                }
            }
        }

        private void DataGridView2_SelectionChanged(object sender, EventArgs e)
        {
           
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void ComboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            switch (comboBox1.Text.ToString())
            {
                case "Все":
                    dataGridView2.DataSource = SqlCon("select * from viewedtzav where Отгрузка is null");

                    break;
                default:
                    dataGridView2.DataSource = SqlCon("select * from viewedtzav where Дилер='" + comboBox1.Text + "' and Отгрузка is null");

                    break;
            }
        }

        private void AddOtgryz_Deactivate(object sender, EventArgs e)
        {

        }
    }
}
