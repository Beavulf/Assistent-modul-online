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
    public partial class EditSotr : Form
    {
        public EditSotr()
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
            //try
            //{
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = sqlstring;
                cmd.ExecuteNonQuery();
                con.Close();
                con.Dispose();
                drg.DataSource = SqlCon("select * from " + table);
                drg.Refresh();
        //}
        //    catch
        //    {
        //        MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        //    }
}
        private void EditSotr_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = SqlCon("select * from Сотрудники");
            dataGridView3.DataSource = SqlCon("select * from viewerr where MONTH(Дата)='"+dateTimePicker2.Value.Month+"'");
            dataGridView4.DataSource = SqlCon("select id_sotr, ФИО from Сотрудники");
            ViewTree(dateTimePicker3.Value.Month);
        }

        private void MaterialRaisedButton4_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;
                // получаем выбранный файл          
                string filename = openFileDialog1.FileName;
                pictureBox1.Image = Image.FromFile(filename);
                InsUpd("update Сотрудники set Фото='" + filename + "' where id_sotr=" + dataGridView1[0, dataGridView1.CurrentRow.Index].Value, "Сотрудники", dataGridView1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }                     
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                string str = dataGridView1[2, dataGridView1.CurrentRow.Index].Value.ToString();
                if (str != "" && str != null)
                    pictureBox1.Image = Image.FromFile(str);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {
            if(materialSingleLineTextField1.Text!=null && materialSingleLineTextField1.Text!="")
            InsUpd("insert into Сотрудники(ФИО) values('"+materialSingleLineTextField1.Text+"')","Сотрудники",dataGridView1);
            else MessageBox.Show("Корректно заполните поля", "Сообщение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            materialSingleLineTextField1.Text = "";
        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Удалить сотрудника: "+ dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString(), "Внимание",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question);
            if (result == DialogResult.Yes)              
            InsUpd("delete Сотрудники where id_sotr="+ dataGridView1[0, dataGridView1.CurrentRow.Index].Value, "Сотрудники",dataGridView1);
        }

        private void MaterialRaisedButton3_Click(object sender, EventArgs e)
        {
            panelEdtM.Visible = true;
            materialRaisedButton7.Visible = false;
            richTextBox1.Text = "";
            materialSingleLineTextField2.Text = "";
            if (materialRaisedButton6.Enabled == true)
                materialRaisedButton6.Enabled = false;
            if (materialRaisedButton5.Enabled == true)
                materialRaisedButton5.Enabled = false;
        }

        private void MaterialRaisedButton10_Click(object sender, EventArgs e)
        {
            try
            {
                InsUpd("insert into Ошибки (id_sotr,Дата, Примечание,Заявка) values(" + dataGridView4[0, dataGridView4.CurrentRow.Index].Value + ",'" + dateTimePicker1.Value.ToShortDateString() + "','" + richTextBox1.Text + "','" + materialSingleLineTextField2.Text + "')", "viewerr where MONTH(Дата)='" + dateTimePicker2.Value.Month + "'", dataGridView3);
            }
            catch { MessageBox.Show("Корректно заполните поля", "Сообщение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }
            richTextBox1.Text = ""; 
            materialSingleLineTextField2.Text = "";
            panelEdtM.Visible = false;
            if (materialRaisedButton6.Enabled == false)
                materialRaisedButton6.Enabled = true;
            if (materialRaisedButton5.Enabled == false)
                materialRaisedButton5.Enabled = true;
            if (materialRaisedButton3.Enabled == false)
                materialRaisedButton3.Enabled = true;
            ViewTree(dateTimePicker3.Value.Month);
        }

        private void MaterialFlatButton2_Click(object sender, EventArgs e)
        {
            if (materialRaisedButton6.Enabled == false)
                materialRaisedButton6.Enabled = true;
            if (materialRaisedButton5.Enabled == false)
                materialRaisedButton5.Enabled = true;
            if (materialRaisedButton3.Enabled == false)
                materialRaisedButton3.Enabled = true;
            richTextBox1.Text = "";
            materialSingleLineTextField2.Text = "";
            panelEdtM.Visible = false;
        }

        private void MaterialRaisedButton5_Click(object sender, EventArgs e)
        {
            try
            {
                var result = MessageBox.Show("Удалить ошибку?", "Внимание",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                    InsUpd("delete Ошибки where id_err="+ dataGridView3[0, dataGridView3.CurrentRow.Index].Value, "viewerr where MONTH(Дата)='" + dateTimePicker2.Value.Month + "'", dataGridView3);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            ViewTree(dateTimePicker3.Value.Month);
        }
        public void ViewTree(int month)
        {
            treeView1.Nodes.Clear();
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            DataTable DT = new DataTable();
            DataTable DT1 = new DataTable();
            SqlCommand cmd = new SqlCommand("select * from viewerr where MONTH(Дата)='" + month + "'", con);
            SqlCommand cmdsotr = new SqlCommand("select id_sotr,ФИО from Сотрудники",con);
            using (SqlDataReader dr1 = cmdsotr.ExecuteReader())
            {
                DT.Load(dr1);
            }
            using (SqlDataReader dr = cmd.ExecuteReader())
            {
                DT1.Load(dr);
            }
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                TreeNode glnode = new TreeNode(DT.Rows[i][1].ToString());//
                for (int j = 0; j < DT1.Rows.Count; j++)
                {
                    if (Convert.ToInt32(DT.Rows[i][0]) == Convert.ToInt32(DT1.Rows[j][5]))
                    {
                        glnode.Nodes.Add(new TreeNode("" + DT1.Rows[j][2].ToString() + ", заявка:" + DT1.Rows[j][3]));
                    }
                }
                treeView1.Nodes.Add(glnode);//
            }
            con.Close();
            con.Dispose();
        }
        private void MaterialRaisedButton6_Click(object sender, EventArgs e)
        {
            if (materialRaisedButton5.Enabled==true)
            {
                materialRaisedButton5.Enabled = false;
            }
            if (materialRaisedButton3.Enabled == true)
            {
                materialRaisedButton3.Enabled = false;
            }
            if (materialRaisedButton7.Visible == false)
            {
                materialRaisedButton7.Visible = true;
            }
            panelEdtM.Visible = true;
            richTextBox1.Text = dataGridView3[4, dataGridView3.CurrentRow.Index].Value.ToString();
            materialSingleLineTextField2.Text = dataGridView3[3, dataGridView3.CurrentRow.Index].Value.ToString();
            dateTimePicker1.Text = dataGridView3[2, dataGridView3.CurrentRow.Index].Value.ToString();
            for(int i =0; i<dataGridView4.RowCount;i++)
            {
                if (dataGridView4.Rows[i].Cells[1].Value.ToString()== dataGridView3[1, dataGridView3.CurrentRow.Index].Value.ToString())
                {
                    dataGridView4.ClearSelection();
                    dataGridView4.Rows[i].Selected = true;
                }
            }
            ViewTree(dateTimePicker3.Value.Month);            
        }

        private void MaterialFlatButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void MaterialRaisedButton7_Click(object sender, EventArgs e)
        {
            InsUpd("update Ошибки set Дата='"+dateTimePicker1.Value+"', Примечание='"+richTextBox1.Text+"', Заявка='"+materialSingleLineTextField2.Text+"' where id_err="+ dataGridView3[0, dataGridView3.CurrentRow.Index].Value, "viewerr where MONTH(Дата)='" + dateTimePicker2.Value.Month + "'", dataGridView3);
            materialRaisedButton7.Visible = false;
            panelEdtM.Visible = false;
            if (materialRaisedButton6.Enabled == false)
                materialRaisedButton6.Enabled = true;
            if (materialRaisedButton5.Enabled == false)
                materialRaisedButton5.Enabled = true;
            if (materialRaisedButton3.Enabled == false)
                materialRaisedButton3.Enabled = true;
        }

        private void DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dataGridView3.DataSource = SqlCon("select * from viewerr where MONTH(Дата)='" + dateTimePicker2.Value.Month + "'");
        }

        private void DateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            ViewTree(dateTimePicker3.Value.Month);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
