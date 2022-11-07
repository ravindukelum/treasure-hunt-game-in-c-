using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Game
{
    public partial class Form1 : Form
    {
        string level;
        public Form1()
        {
            InitializeComponent();
        }

        private void linkHelp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Help h = new Help();
            h.Show();
            this.Hide();
        }

        private void btnstart_Click(object sender, EventArgs e)
        {
          
           
            if (btneasy.Checked)
            {
                level = "Easy";
            }
            if (btnmedium.Checked)
            {
                level = "Medium";
            }
            if (btnHard.Checked)
            {
                level = "Hard";
            }
            GamePage g = new GamePage(txboxUserName.Text,level);
            g.Show();
            this.Hide();
        }
    }
}
