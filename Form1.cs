using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SwCSharpAddinByStanley
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.textBox1.Text = ConfigurationManipulate.GetConfigValue("图号前缀");
            if(ConfigurationManipulate.GetConfigValue("随机颜色") == "True")
            {
                this.checkBox1.Checked = true;
            }
            else
            {
                this.checkBox1.Checked = false;
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConfigurationManipulate.SetConfigValue("图号前缀", this.textBox1.Text);
            ConfigurationManipulate.SetConfigValue("随机颜色", this.checkBox1.Checked.ToString());
            this.Close();
        }
    }
}
