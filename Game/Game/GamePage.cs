using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Game
{
    public partial class GamePage : Form
    {
        Workbook wb;
        Worksheet ws;
        int val = 1;
        string CellValue;
        string name,level;
        int chance;
        string vv = "one";
        public GamePage(string name,string level)
        {
            InitializeComponent();
            this.name = name;
            this.level = level;
            if(level=="Easy")
            {
                chance = 40;
                lblChance.Text =Convert.ToString(chance);

            }
            else if(level == "Medium")
            {
                chance = 35;
                lblChance.Text = Convert.ToString(chance);

            }
            else if (level == "Hard")
            {
                chance = 30;
                lblChance.Text = Convert.ToString(chance);

            }
            lblname.Text = name;
            string filePath = "C:\\Users\\Ravindu\\Desktop\\Project\\gamepp.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
           
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];
        }

      
        public void ReadExcel(int i,int y)
        {
            /*
            string filePath = "C:\\Users\\Ravindu\\Desktop\\Project\\gamepp.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];
            */
            Range cell = ws.Cells[i, y];
            CellValue =Convert.ToString(cell.Value);
            
           //MessageBox.Show(CellValue);
           // vv++;
        }
        public void extra()
        {
            if(level=="Easy")
            {
                MessageBox.Show("Your Have more 10 Chancess");
                chance = 10;
                lblChance.Text =Convert.ToString(10);
            }
            else if(level == "Medium")
            {
                MessageBox.Show("Your Have more 6 Chancess");
                chance = 6;
                lblChance.Text = Convert.ToString(6);
            }
            else if (level == "Hard")
            {
                MessageBox.Show("Your Have more 3 Chancess");
                chance = 3;
                lblChance.Text = Convert.ToString(3);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance==0)
            {
              
                if(vv=="one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }
                
                
            }
            
            ReadExcel(1, 1);
            if (CellValue == "0")
            {
                button1.BackColor = Color.Red;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 2);
            if (CellValue == "0")
            {
                button2.BackColor = Color.Red;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 3);
            if (CellValue == "0")
            {
                button3.BackColor = Color.Red;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 5);
            if (CellValue == "0")
            {
                button5.BackColor = Color.Red;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 6);
            if (CellValue == "0")
            {
                button6.BackColor = Color.Red;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 7);
            if (CellValue == "0")
            {
                button7.BackColor = Color.Red;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 8);
            if (CellValue == "0")
            {
                button8.BackColor = Color.Red;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 9);
            if (CellValue == "0")
            {
                button9.BackColor = Color.Red;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 10);
            if (CellValue == "0")
            {
                button10.BackColor = Color.Red;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 1);
            if (CellValue == "0")
            {
                button11.BackColor = Color.Red;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 2);
            if (CellValue == "0")
            {
                button12.BackColor = Color.Red;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 3);
            if (val >= 3)
            {
                
                  if (CellValue == "3")
                    {
                    button13.BackColor = Color.Green;
                    val = 4;
                }
              

            }
            else
            {

                button13.BackColor = Color.Red;

            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 4);
            if (CellValue == "0")
            {
                button14.BackColor = Color.Red;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2,5);
            if (val >= 5)
            {

                if (CellValue == "5")
                {
                    button15.BackColor = Color.Green;
                    val = 6;
                }


            }
            else
            {

                button15.BackColor = Color.Red;

            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 6);
            if (CellValue == "0")
            {
                button16.BackColor = Color.Red;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 7);
            if (CellValue == "0")
            {
                button17.BackColor = Color.Red;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 8);
            if (CellValue == "0")
            {
                button18.BackColor = Color.Red;
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2,9);
            if (val >= 12)
            {

                if (CellValue == "ggg")
                {
                    button19.BackColor = Color.Green;
                    MessageBox.Show("Congratulations, You Found The Treashure");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                    val = 14;
                }


            }
            else
            {

                button19.BackColor = Color.Red;

            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(2, 10);
            if (CellValue == "0")
            {
                button20.BackColor = Color.Red;
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3,10);
            if (val >= 12)
            {

                if (CellValue == "12")
                {
                    button30.BackColor = Color.Green;
                    val = 13;
                }


            }
            else
            {

                button30.BackColor = Color.Red;

            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 9);
            if (CellValue == "0")
            {
                button29.BackColor = Color.Red;
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 8);
            if (CellValue == "0")
            {
                button28.BackColor = Color.Red;
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 7);
            if (CellValue == "0")
            {
                button27.BackColor = Color.Red;
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 6);
            if (CellValue == "0")
            {
                button26.BackColor = Color.Red;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 5);
            if (val >= 6)
            {

                if (CellValue == "6")
                {
                    button25.BackColor = Color.Green;
                    val = 7;
                }


            }
            else
            {

                button25.BackColor = Color.Red;

            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 4);
            if (CellValue == "0")
            {
                button24.BackColor = Color.Red;
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 3);
            if (CellValue == "0")
            {
                button23.BackColor = Color.Red;
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 2);
            if (val >= 2)
            {
                if (CellValue == "2")
                {
                    button22.BackColor = Color.Green;
                    val = 3;
                }

            }
            else
            {
                
                 button22.BackColor = Color.Red;
               
            }

        }

        private void button21_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(3, 1);
            if (CellValue == "0")
            {
                button21.BackColor = Color.Red;
            }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 10);
            if (CellValue == "0")
            {
                button50.BackColor = Color.Red;
            }
        }

        private void button49_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 9);
            if (CellValue == "0")
            {
                button49.BackColor = Color.Red;
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5,8);
            if (val >= 10)
            {

                if (CellValue == "10")
                {
                    button48.BackColor = Color.Green;
                    val = 11;
                }


            }
            else
            {

                button48.BackColor = Color.Red;

            }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5,7);
            if (val >= 9)
            {

                if (CellValue == "9")
                {
                    button47.BackColor = Color.Green;
                    val = 10;
                }


            }
            else
            {

                button47.BackColor = Color.Red;

            }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 6);
            if (CellValue == "0")
            {
                button46.BackColor = Color.Red;
            }
        }

        private void button45_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 5);
            if (CellValue == "0")
            {
                button45.BackColor = Color.Red;
            }
        }

        private void button44_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 4);
            if (CellValue == "0")
            {
                button44.BackColor = Color.Red;
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 3);
            if (CellValue == "0")
            {
                button43.BackColor = Color.Red;
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 2);
            if (CellValue == "0")
            {
                button42.BackColor = Color.Red;
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(5, 1);
            if (CellValue == "0")
            {
                button41.BackColor = Color.Red;
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 10);
            if (CellValue == "0")
            {
                button40.BackColor = Color.Red;
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4,9);
            if (val >= 11)
            {

                if (CellValue == "11")
                {
                    button39.BackColor = Color.Green;
                    val = 12;
                }


            }
            else
            {

                button39.BackColor = Color.Red;

            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 8);
            if (CellValue == "0")
            {
                button38.BackColor = Color.Red;
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 7);
            if (val >= 8)
            {

                if (CellValue == "8")
                {
                    button37.BackColor = Color.Green;
                    val = 9;
                }


            }
            else
            {

                button37.BackColor = Color.Red;

            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4,6);
            if (val >= 7)
            {

                if (CellValue == "7")
                {
                    button36.BackColor = Color.Green;
                    val = 8;
                }


            }
            else
            {

                button36.BackColor = Color.Red;

            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 5);
            if (CellValue == "0")
            {
                button35.BackColor = Color.Red;
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 4);
            if (CellValue == "0")
            {
                button34.BackColor = Color.Red;
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 3);
            if (CellValue == "0")
            {
                button33.BackColor = Color.Red;
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 2);
            if (CellValue == "0")
            {
                button32.BackColor = Color.Red;
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(4, 1);
            if (CellValue == "1")
            {
                
                if (val == 1)
                {
                    button31.BackColor = Color.Green;
                    val = 2;
                }
              
                
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            chance--;
            lblChance.Text = Convert.ToString(chance);
            if (chance == 0)
            {

                if (vv == "one")
                {
                    vv = "Two";
                    extra();
                }
                else
                {
                    MessageBox.Show("Game Over");
                    Form1 f = new Form1();
                    f.Show();
                    this.Hide();
                }


            }
            ReadExcel(1, 4);
            if (val >= 4)
            {

                if (CellValue == "4")
                {
                    button4.BackColor = Color.Green;
                    val = 5;
                }


            }
            else
            {

                button4.BackColor = Color.Red;

            }
        }

        private void GamePage_Load(object sender, EventArgs e)
        {

        }
    }
}
