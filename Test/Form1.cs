using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        KellPrinter.DataReporter report;

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable data = new DataTable("Test Report");
            data.Columns.Add(new DataColumn("Name", typeof(string)));
            data.Columns.Add(new DataColumn("Home", typeof(string)));
            data.Columns.Add(new DataColumn("Quantity", typeof(int)));
            data.Columns.Add(new DataColumn("Status", typeof(bool)));
            Random ran = new Random();
            for (int i = 0; i < 10; i++)
            {
                DataRow row = data.NewRow();
                row.ItemArray = new object[] { "bill" + ran.Next(10), "东莞市东城区", ran.Next(100), ran.Next(2) == 1 ? true : false };
                data.Rows.Add(row);
            }
            string[] header = { "序号", "名称", "住址", "数量", "状态" };
            string[] bottom = { "制表", "核准" };
            Dictionary<int, SortOrder> sort = new Dictionary<int, SortOrder>();
            sort.Add(0, SortOrder.None);//Descending
            data = KellPrinter.DataReporter.Sort(data, sort);
            KellPrinter.PrintArgs args = new KellPrinter.PrintArgs("Test", "Report", header, bottom);
            report = new KellPrinter.DataReporter(data, args);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            report.PrintReport();
            //PrintPreviewDialog prt = report.PrintReport();            
            //PrintDialog pd = new PrintDialog();
            //pd.Document = prt.Document;
            //pd.UseEXDialog = true;
            //if (pd.ShowDialog() == DialogResult.OK)
            //    prt.Document.Print();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Exception ex = report.SaveAsExcel(null);
            if (ex == null)
                MessageBox.Show("导出成功！");
            else
                MessageBox.Show("导出失败：" + ex.Message);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PrintPreviewDialog p = report.PreviewPrintReport();
            panel1.Controls.Add(p);
            p.Show();
            panel1.Refresh();
        }
    }
}
