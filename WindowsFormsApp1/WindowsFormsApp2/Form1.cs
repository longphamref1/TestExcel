using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CreateNewFile();
        }

        
        public void OpenFile()
        {
            Excel excel = new Excel(@"C:\VAS_CDR_report\19.02.11\VAS.CDR.report.190211.xls", 1);
            MessageBox.Show(excel.ReadCell(20, 0));
        }

        public void WriteData()
        {
            Excel excel = new Excel(@"C:\VAS_CDR_report\19.02.11\VAS.CDR.report.190211.xls", 1);
            excel.WriteToCell(0, 10, "Test Write");
            excel.Save();
            excel.Close();
            excel.SaveAs(@"C:\VAS_CDR_report\19.02.11\VAS.CDR.report.1902112.xls");
            excel.Close();
        }

        public void CreateNewFile()
        {
            Excel ex = new Excel();
            ex.CreateNewFile();
            ex.SaveAs(@"C:\VAS_CDR_report\19.02.11\Test.xlsx");
            ex.Close();
        }
    }
}
