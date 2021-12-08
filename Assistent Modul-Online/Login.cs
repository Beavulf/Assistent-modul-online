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
    public partial class Login : MaterialForm
    {
        public Login()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue800, Primary.Blue400, Accent.LightBlue200, TextShade.WHITE);
        }

        private void MaterialRaisedButton1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == Properties.Resources.PasswordZav)
            {
                this.Close();
                ZavSklad zvs = new ZavSklad();
                zvs.ShowDialog();              
            }
            else { MessageBox.Show("Не правильный пароль", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }
            
        }

        private void MaterialRaisedButton2_Click(object sender, EventArgs e)
        {
            
        }
    }
}
