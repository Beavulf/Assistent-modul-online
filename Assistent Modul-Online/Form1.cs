using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MaterialSkin;
using MaterialSkin.Controls;
using Microsoft.VisualBasic;

namespace Assistent_Modul_Online
{
    public partial class Form1 : MaterialForm
    {
        public Form1()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue800, Primary.Blue400, Accent.LightBlue200, TextShade.WHITE);
        }

        string connstring = @"Server=.\SQLEXPRESS;Database=modul-online;Trusted_Connection=true;";
        

        public void InsUpd(string sqlstring, string table, DataGridView drg)
        { //выполнение команды (добав. удал, измн)
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            //try
            //{
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = sqlstring;
                cmd.ExecuteNonQuery();
                con.Close();
                drg.DataSource = SqlCon("select * from " + table);
                drg.Refresh();
            //}
            //catch
            //{
            //    MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            //}
        }

        public void PaintGrid()
        {//окрашиваем строки
            foreach (DataGridViewRow row in dataGridView1.Rows)
                if (Convert.ToInt32(row.Cells[5].Value) == 1)
                {
                    row.DefaultCellStyle.BackColor = Color.LightGreen;
                }
        }


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
            //}
            //catch
            //{
            //    MessageBox.Show("Ошибка подключения к серверу");
            //}
            return DT;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = SqlCon("select * from viewzav where MONTH(Дата)=MONTH(FORMAT(getdate(),'yyyy-MM-dd'))");
            dataGridView2.DataSource = SqlCon("select * from viewvrabote where MONTH(Дата)=MONTH(FORMAT(getdate(),'yyyy-MM-dd'))");
            SqlDataAdapter vid1 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable fam1 = new DataTable();
            vid1.Fill(fam1);
            comboBox1.DataSource = fam1;
            comboBox1.DisplayMember = "ФИО";
            comboBox1.ValueMember = "id_sotr";

            SqlDataAdapter vid2 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable fam2 = new DataTable();
            vid2.Fill(fam2);
            comboBox2.DataSource = fam2;
            comboBox2.DisplayMember = "ФИО";
            comboBox2.ValueMember = "id_sotr";

            SqlDataAdapter vid5 = new SqlDataAdapter("select * from Другие_работы", connstring);
            DataTable rab = new DataTable();
            vid5.Fill(rab);
            comboBox3.DataSource = rab;
            comboBox3.DisplayMember = "Название";
            comboBox3.ValueMember = "id_rab";

            SqlDataAdapter vid6 = new SqlDataAdapter("select * from Диллеры", connstring);
            DataTable dil2 = new DataTable();
            vid6.Fill(dil2);
            comboBox5.DataSource = dil2;
            comboBox5.DisplayMember = "Наименование";
            comboBox5.ValueMember = "id_dil";

            SqlDataAdapter vid3 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable fam3 = new DataTable();
            vid3.Fill(fam3);
            comboBox4.DataSource = fam3;
            comboBox4.DisplayMember = "ФИО";
            comboBox4.ValueMember = "id_sotr";
            
        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        {
            if ((materialSingleLineTextField1.Text != "" && materialSingleLineTextField1.Text != null) && comboBox5.Text != "Все")
            {
                string insert = "insert into Заявка (Номер, id_sotr, Время_начала, Дата_сборки, Статус, Строки, Штуки, id_dil) values ('" + materialSingleLineTextField1.Text + "','" + comboBox1.SelectedValue + "', FORMAT(getdate(),'hh:mm:ss'),FORMAT(getdate(),'yyyy-MM-dd'),'0','" + materialSingleLineTextField3.Text + "','" + materialSingleLineTextField2.Text + "','" + comboBox5.SelectedValue + "')";
                InsUpd(insert, "viewzav where MONTH(Дата)=MONTH(FORMAT(getdate(),'yyyy-MM-dd'))", dataGridView1);
                materialSingleLineTextField1.Text = "";
                materialSingleLineTextField2.Text = "";
                materialSingleLineTextField3.Text = "";
            }
            else MessageBox.Show("Корректно заполните поля", "Сообщение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {

            if ((materialSingleLineTextField1.Text != "" && materialSingleLineTextField1.Text != null))
            {
                SqlConnection con = new SqlConnection(connstring);
                con.Open();
                try
                {
                    SqlCommand cmd = con.CreateCommand();
                    string update = "exec CloseZaiav " + materialSingleLineTextField1.Text + "";
                    cmd.CommandText = update;
                    int a = cmd.ExecuteNonQuery();
                    con.Close();
                    con.Dispose();
                    dataGridView1.DataSource = SqlCon("select * from viewzav where MONTH(Дата)=MONTH(FORMAT(getdate(),'yyyy-MM-dd'))");
                    dataGridView1.Refresh();
                    materialSingleLineTextField1.Text = "";
                    if (a > 0) MessageBox.Show("Заявка успешна закрыта", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    else MessageBox.Show("Заявка не найдена", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);                  
                }
                catch { MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }
                }
            else MessageBox.Show("Введите номер заявки необходимой для закрытия", "Сообщение", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1); 
        }

        private void DataGridView1_Paint(object sender, PaintEventArgs e)
        {
            PaintGrid();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            
        }

        private void ComboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void MaterialRaisedButton3_Click(object sender, EventArgs e)
        {
            string ins = " if (select count(*) from Вработе where id_sotr='"+comboBox2.SelectedValue+"' and id_rab='"+ comboBox3.SelectedValue + "' and Отработал is null and Дата=FORMAT(getdate(),'yyyy-MM-dd'))>0 print('') else insert into Вработе (id_sotr, id_rab, Дата, Время) values ('" + comboBox2.SelectedValue + "','" + comboBox3.SelectedValue + "',FORMAT(getdate(),'yyyy-MM-dd'),FORMAT(getdate(),'hh:mm:ss'))";

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {            
               if (dataGridView1.Rows[i].Cells[3].Value != null)
                   if (dataGridView1.Rows[i].Cells[3].Value.ToString().Contains(comboBox2.Text) && (dataGridView1.Rows[i].Cells[4].Value.ToString()==""))
                   {
                            MessageBox.Show("Нельзя повторно отправить сотрудника на другую работу");
                   }
               else { InsUpd(ins, "viewvrabote where Дата=FORMAT(getdate(),'yyyy-MM-dd')", dataGridView2); break; };
            }
            
                      
        }

        private void MaterialRaisedButton4_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            try
            {
                SqlCommand cmd = con.CreateCommand();
                string update = "update Вработе set Время_окончания=FORMAT(getdate(),'hh:mm:ss'), Отработал=DATEDIFF(MINUTE,Время,FORMAT(getdate(),'hh:mm:ss')) where Дата=FORMAT(getdate(),'yyyy-MM-dd') and Время_окончания is null and id_sotr=" + comboBox2.SelectedValue;
                cmd.CommandText = update;
                int a = cmd.ExecuteNonQuery();
                con.Close();
                //con.Dispose();
                dataGridView2.DataSource = SqlCon("select * from viewvrabote where Дата=FORMAT(getdate(),'yyyy-MM-dd')");
                dataGridView2.Refresh();
                if (a > 0) MessageBox.Show("Работа окончена", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                else MessageBox.Show("Сотрудник в работе не найден", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
            catch (Exception ex )
            {
               
                MessageBox.Show("Ошибка подключения к серверу: "+ex.Message, "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }           
        }

        private void MaterialSingleLineTextField3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void MaterialSingleLineTextField2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void MaterialSingleLineTextField1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsDigit(e.KeyChar) && e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            

        }


        private void MaterialSingleLineTextField4_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void GroupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            Report rep = new Report();
            rep.ShowDialog();
        }

        private void MaterialRaisedButton5_Click(object sender, EventArgs e)
        {
            Login lg = new Login();
            lg.Show();
        }

        private void MaterialRaisedButton7_Click(object sender, EventArgs e)
        {
            AddOtgryz add = new AddOtgryz();
            add.ShowDialog();
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewvrabote where Дата='" + dateTimePicker1.Value.ToShortDateString() + "'");
        }

        private void ComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox4.Text.ToString())
            {
                case "Все":
                    dataGridView2.DataSource = SqlCon("select * from viewvrabote where Дата='" + dateTimePicker1.Value.ToShortDateString() + "'");

                    break;
                default:
                    dataGridView2.DataSource = SqlCon("select * from viewvrabote where Дата='" + dateTimePicker1.Value.ToShortDateString() + "' and ФИО='" + comboBox4.Text + "'");

                    break;
            }
        }

        private void MaterialFlatButton1_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select * from viewvrabote where Дата='" + dateTimePicker1.Value.ToShortDateString() + "'");
        }
    }
}
