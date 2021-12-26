using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace трпо
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public Form2 (Form1 f)
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        public void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                if (dataGridView2[0,i].FormattedValue.ToString() == SearchBox.Text)
                {
                    dataGridView2.Rows.RemoveAt(i--);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            int c;

            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                string input = textBox1.Text;
                var date = DateTime.ParseExact(input, "dd-MM-yyyy", null, DateTimeStyles.AssumeUniversal);
                //DateTime dateTime = Convert.ToDateTime(DateTime.FromOADate(double.Parse(SearchBox.Text)));
                DateTime dat = date.Date;

                System.TimeSpan diff = Convert.ToDateTime(dat) -
                    Convert.ToDateTime(main.dataGridView1.Rows[i].Cells[3].Value);

                c = diff.Days;

                if (c < 0)
                {
                    dataGridView2.Rows[i].Cells[1].Value = 0; 
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
