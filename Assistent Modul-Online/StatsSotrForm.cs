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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace Assistent_Modul_Online
{
    public partial class StatsSotrForm : Form
    {
        public StatsSotrForm()
        {
            InitializeComponent();

        }
        string connstring = @"Server=.\SQLEXPRESS;Database=modul-online;Trusted_Connection=true;";
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
                MessageBox.Show("Ошибка подключения к серверу", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
            return DT;
        }

        public Excel.Application application = null;
        public Excel.Workbook workbook = null;
        public Excel.Workbooks workbooks = null;
        public Excel.Sheets worksheets = null;
        public Excel.Worksheet worksheet = null;


            
        

        private void StatsSotrForm_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = SqlCon("select * from viewreport where MONTH(Дата)='"+dateTimePicker1.Value.Month+"'");

            SqlDataAdapter vid1 = new SqlDataAdapter("select * from Сотрудники", connstring);
            DataTable fam3 = new DataTable();
            vid1.Fill(fam3);
            comboBox1.DataSource = fam3;
            comboBox1.DisplayMember = "ФИО";
            comboBox1.ValueMember = "id_sotr";
        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(connstring);
            con.Open();
            if (materialCheckBox1.Checked == true)
            {
                string str = dateTimePicker1.Value.Month.ToString();
                SqlCommand cmd = new SqlCommand("CalcReportFullMonth", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@sotr", SqlDbType.Int).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@data", SqlDbType.VarChar, 3).Value = str;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@atime", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@psborki", SqlDbType.Float));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@atime"].Direction = ParameterDirection.Output;
                cmd.Parameters["@psborki"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label3.Text = cmd.Parameters["@count"].Value.ToString();
                label7.Text = cmd.Parameters["@acount"].Value.ToString() + " строк(и)";
                label9.Text = cmd.Parameters["@atime"].Value.ToString()+" минут(ы)";
                label11.Text = cmd.Parameters["@psborki"].Value.ToString() + " %";
                string report = "select * from viewreport where ФИО='" + comboBox1.Text + "' and MONTH(Дата) =" + str;
                dataGridView1.DataSource = SqlCon(report);
            }
            else
            {
                SqlCommand cmd = new SqlCommand("CalcReportFull", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@sotr", SqlDbType.Int).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@data", SqlDbType.Date).Value = dateTimePicker1.Value;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@atime", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@psborki", SqlDbType.Float));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@atime"].Direction = ParameterDirection.Output;
                cmd.Parameters["@psborki"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label3.Text = cmd.Parameters["@count"].Value.ToString();
                label7.Text = cmd.Parameters["@acount"].Value.ToString()+" строк(и)";
                label9.Text = cmd.Parameters["@atime"].Value.ToString()+ " минут(ы)";
                label11.Text = cmd.Parameters["@psborki"].Value.ToString()+ " %";
                string report1 = "select * from viewreport where Дата='" + dateTimePicker1.Value + "' and ФИО='" + comboBox1.Text + "'";
                dataGridView1.DataSource = SqlCon(report1);
            }

            if (materialRadioButton2.Checked == true)
            {
                SqlCommand cmd = new SqlCommand("CalcReportFullBetween", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@sotr", SqlDbType.Int).Value = comboBox1.SelectedValue;
                cmd.Parameters.Add("@data1", SqlDbType.Date).Value = dateTimePicker2.Value;
                cmd.Parameters.Add("@data2", SqlDbType.Date).Value = dateTimePicker3.Value;
                cmd.Parameters.Add(new SqlParameter("@count", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@acount", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@atime", SqlDbType.Int));
                cmd.Parameters.Add(new SqlParameter("@psborki", SqlDbType.Float));
                cmd.Parameters["@count"].Direction = ParameterDirection.Output;
                cmd.Parameters["@acount"].Direction = ParameterDirection.Output;
                cmd.Parameters["@atime"].Direction = ParameterDirection.Output;
                cmd.Parameters["@psborki"].Direction = ParameterDirection.Output;
                cmd.ExecuteNonQuery();
                label3.Text = cmd.Parameters["@count"].Value.ToString();
                label7.Text = cmd.Parameters["@acount"].Value.ToString() + " строк(и)";
                label9.Text = cmd.Parameters["@atime"].Value.ToString() + " минут(ы)";
                label11.Text = cmd.Parameters["@psborki"].Value.ToString() + " %";
                string report2 = "select * from viewreport where Дата between '" + dateTimePicker2.Value + "' and '"+dateTimePicker3.Value+"' and ФИО='" + comboBox1.Text + "'";
                dataGridView1.DataSource = SqlCon(report2);
            }

            con.Close();
            con.Dispose();
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {                   
            chart1.Series[0].IsValueShownAsLabel = true;
            chart1.Series[0].Font = new Font("Microsoft Sans Serif", 11);
            chart1.Titles.Clear();
            chart1.Series["Series1"].Points.Clear();
            chart1.Titles.Add("График");
            chart1.Titles[0].Font = new Font("Microsoft Sans Serif", 11);
            chart1.Series["Series1"].LegendText = "График количества строк по датам";
            chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            dataGridView2.DataSource = SqlCon("select Дата_сборки, SUM(Строки) as Строки from Заявка where MONTH(Дата_сборки) = '"+dateTimePicker1.Value.Month+"' group by Дата_сборки");
            
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {               
                chart1.Series["Series1"].Points.AddXY(Convert.ToDateTime(row.Cells[0].Value).ToShortDateString(), row.Cells[1].Value);
                             
            }
        }

        private void Label7_Click(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void MaterialFlatButton1_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select Дата_сборки, SUM(Строки) as Строки from Заявка where MONTH(Дата_сборки) = '" + dateTimePicker1.Value.Month + "' group by Дата_сборки");
            try
            {
                application = new Excel.Application();
                application.SheetsInNewWorkbook = 1;
                workbooks = application.Workbooks;
                workbook = workbooks.Add();
                worksheets = application.Sheets;
                worksheet = worksheets.Item[1];
                worksheet.Name = "График строк за " + dateTimePicker1.Value.Month.ToString() + " " + dateTimePicker1.Value.Year.ToString() + "г";
                application.Visible = true;
                worksheet.Cells.EntireColumn.AutoFit();
                worksheet.Cells.EntireRow.AutoFit();
                worksheet.Cells[1, 1] = "Отчет за ";
                worksheet.Cells[1, 1].Font.Bold = true;
                worksheet.Cells[1, 2].Font.Bold = true;
                worksheet.Cells[1, 2] = dateTimePicker1.Value.ToLongDateString();
                worksheet.Cells.Font.Size = 16;

                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    {
                        worksheet.Cells.EntireColumn.AutoFit();
                        worksheet.Cells.EntireRow.AutoFit();
                        worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        
                    }
                }

                worksheet.Cells[dataGridView2.RowCount + 2, 1].Value = "Итого: ";
                worksheet.Cells[dataGridView2.RowCount + 2, 1].Font.Bold = true;
                worksheet.Cells[dataGridView2.RowCount + 2, 2].Value = "ss";
                worksheet.Cells[dataGridView2.RowCount + 2, 2].FormulaLocal = "=СУММ(B2:B"+ (dataGridView2.RowCount+1) + ")";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                application.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(worksheets);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(application);
            }
        }

        private void MaterialRaisedButton3_Click(object sender, EventArgs e)
        {
            chart1.Series[0].IsValueShownAsLabel = true;
            chart1.Series[0].Font = new Font("Microsoft Sans Serif", 11);
            chart1.Titles.Clear();
            chart1.Series["Series1"].Points.Clear();
            chart1.Titles.Add("График");
            chart1.Titles[0].Font = new Font("Microsoft Sans Serif", 11);
            chart1.Series["Series1"].LegendText = "График количества строк сотрудников";
            chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            dataGridView2.DataSource = SqlCon("select ФИО, SUM(Строки) as Строки from viewzav where MONTH(Дата) = '" + dateTimePicker1.Value.Month + "' group by ФИО");

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                chart1.Series["Series1"].Points.AddXY(Convert.ToString(row.Cells[0].Value), row.Cells[1].Value);

            }
        }

        private void PictureBox3_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = SqlCon("select ФИО, SUM(Строки) as Строки from viewzav where MONTH(Дата) = '" + dateTimePicker1.Value.Month + "' group by ФИО");

            try
            {
                application = new Excel.Application();
                application.SheetsInNewWorkbook = 1;
                workbooks = application.Workbooks;
                workbook = workbooks.Add();
                worksheets = application.Sheets;
                worksheet = worksheets.Item[1];
                worksheet.Name = "График строк за " + dateTimePicker1.Value.Month.ToString() + " " + dateTimePicker1.Value.Year.ToString() + "г";
                application.Visible = true;
                worksheet.Cells.EntireColumn.AutoFit();
                worksheet.Cells.EntireRow.AutoFit();
                worksheet.Cells[1, 1].Font.Bold = true;
                worksheet.Cells[1, 2].Font.Bold = true;
                worksheet.Cells[1, 1] = "Отчет за ";
                worksheet.Cells[1, 2] = dateTimePicker1.Value.ToLongDateString();
                worksheet.Cells.Font.Size = 16;

                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    {
                        worksheet.Cells.EntireColumn.AutoFit();
                        worksheet.Cells.EntireRow.AutoFit();
                        worksheet.Cells[i + 2, j + 1] = String.Format(dataGridView2.Rows[i].Cells[j].Value.ToString());
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                application.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(worksheets);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(application);
            }
        }

        private void MaterialCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            materialRadioButton2.Checked = false;
        }

        private void MaterialRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            materialCheckBox1.Checked = false;
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = SqlCon("select * from viewreport where MONTH(Дата)='" + dateTimePicker1.Value.Month + "'");
        }
    }
   
}
