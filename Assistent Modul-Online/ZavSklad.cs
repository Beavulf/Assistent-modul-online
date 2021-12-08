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
using System.Windows.Media;
using System.Drawing.Drawing2D;
namespace Assistent_Modul_Online
{
    public partial class ZavSklad : MaterialForm
    {
        public ZavSklad()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue800, Primary.Blue400, Accent.LightBlue200, TextShade.WHITE);
            HidePanel();
            Timer timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += new EventHandler(Timer1_Tick);
            timer.Start();
            GetDate(label2);
        }
        public void HidePanel()
        {
            panelSubMenuStats.Visible = false;
            panelSubEditMenu.Visible = false;
        }

        public void HideSubMenu()
        {//включение выключение меню
            if (panelSubMenuStats.Visible == true)
                panelSubMenuStats.Visible = false;
            if (panelSubEditMenu.Visible == true)
                panelSubEditMenu.Visible = false;
        }

        public void ShowMenu(Panel subpanel)
        {//появление меню
            if (subpanel.Visible == false)
            {
                HideSubMenu();
                subpanel.Visible = true;
            }
            else
                subpanel.Visible = false;
        }


        private Form activeForm = null;
        private void openChildForm(Form cildform)
        {//отображение новой формы в окне
            if (activeForm != null)
                activeForm.Close();
            activeForm = cildform;
            cildform.TopLevel = false;
            cildform.FormBorderStyle = FormBorderStyle.None;
            cildform.Dock = DockStyle.Fill;
            grdPanelChieldForm.Controls.Add(cildform);
            grdPanelChieldForm.Tag = cildform;
            cildform.BringToFront();
            cildform.Show();
        }
        private void ZavSklad_Load(object sender, EventArgs e)
        {
            
        }
     
        private void TabPage1_Click(object sender, EventArgs e)
        {

        }

        private void BtnStats_Click(object sender, EventArgs e)
        {
            ShowMenu(panelSubMenuStats);
        }

        private void BtnSotr_Click(object sender, EventArgs e)
        {
            openChildForm(new StatsSotrForm());
        }

        private void BntEdit_Click(object sender, EventArgs e)
        {
            ShowMenu(panelSubEditMenu);
        }

        public void GetDate (Label lbl)
        {
            string data = "";
            int day = DateTime.Now.Day;
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;

            if (day < 10)
            {
                data += "0" + day;
            }
            else
            {
                data += day;
            }
            data += ".";
            if (month < 10)
            {
                data += "0" + month;
            }
            else
            {
                data += month;
            }
            data += ".";
            data += year;
            lbl.Text = data;
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            int h = DateTime.Now.Hour;
            int m = DateTime.Now.Minute;
            int s = DateTime.Now.Second;

            string time = "";
                  
            if (h < 10)
            {
                time += "0" + h;
            }
            else
            {
                time += h;
            }

            time += ":";

            if (m < 10)
            {
                time += "0" + m;
            }
            else
            {
                time += m;
            }

            time += ":";

            if (s < 10)
            {
                time += "0" + s;
            }
            else
            {
                time += s;
            }
            label1.Text = time;
        }

        private void BtnZav_Click(object sender, EventArgs e)
        {
            openChildForm(new StatsZayav());
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            openChildForm(new EditSotr());
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            openChildForm(new EditZayav());
        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            openChildForm(new EditDilRab());
        }
    }
}
