using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;


namespace трпо
{
    public partial class Form1 : Form
    {        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int a = 0;
            int c = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                a += Convert.ToInt32(dataGridView1[6,i].Value);
                c++;
            }
            c--;
            textBox1.Text = a.ToString();
            textBox2.Text = c.ToString();
            int sr = a / c;
            textBox3.Text = sr.ToString();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Deletebutton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int c = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (dataGridView1[5, i].FormattedValue.ToString() == SearchBox.Text)
                {
                    c++;
                }
            }
            textBox4.Text = Convert.ToString(c);
            return;
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {            
            Form2 newForm2 = new Form2();
            newForm2.Owner = this;
            newForm2.Show();

            newForm2.dataGridView2.RowCount = dataGridView1.RowCount;
            string c;
            int C;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                c = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                newForm2.dataGridView2.Rows[i].Cells[0].Value = c;

                C = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value);
                newForm2.dataGridView2.Rows[i].Cells[1].Value = C;

                System.TimeSpan diff = Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].Value) -
                    Convert.ToDateTime(dataGridView1.Rows[i].Cells[3].Value);

                newForm2.dataGridView2.Rows[i].Cells[2].Value = diff.Days;
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            int c = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (Convert.ToInt32(dataGridView1[6, i].Value) > c)
                {
                    c = Convert.ToInt32(dataGridView1[6, i].Value);
                }
            }
            textBox5.Text = Convert.ToString(c);
            return;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            {
                double date;
                string sDate;
                string str;
                int str1;
                int rCnt;
                int cCnt;

                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel (*.XLSX)|*.XLSX";
                opf.ShowDialog();
                DataTable tb = new DataTable();
                string filename = opf.FileName;

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                Microsoft.Office.Interop.Excel.Range ExcelRange;

                ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                    false, 0, true, 1, 0);
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                ExcelRange = ExcelWorkSheet.UsedRange;
                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for (cCnt = 1; cCnt <= 1; cCnt++)
                    {
                        str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                    for (cCnt = 2; cCnt <= 2; cCnt++)
                    {
                        str1 = Convert.ToInt32((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str1;
                    }
                    for (cCnt = 3; cCnt <= 3; cCnt++)
                    {
                        str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                    for (cCnt = 4; cCnt <= 5; cCnt++)
                    {
                        sDate = (ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                        date = double.Parse(sDate);
                        DateTime dateTime = Convert.ToDateTime(DateTime.FromOADate(date));
                        DateTime dat = dateTime.Date;
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = dat;
                    }
                    for (cCnt = 6; cCnt <= 6; cCnt++)
                    {
                        str = Convert.ToString((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                    for (cCnt = 7; cCnt <= 7; cCnt++)
                    {
                        str1 = Convert.ToInt32((ExcelRange.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str1;
                    }
                }
                ExcelWorkBook.Close(true, null, null);
                ExcelApp.Quit();

                releaseObject(ExcelWorkSheet);
                releaseObject(ExcelWorkBook);
                releaseObject(ExcelApp);
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
