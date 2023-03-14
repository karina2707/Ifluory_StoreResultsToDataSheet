using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace LbFiveInterf
{
    public partial class Form1 : Form
    {
        private int compCount;
        public Form1()
        {
            InitializeComponent();
            compCount = Program.GetCompanyCount();
            numericUpDownTo.Maximum = compCount;
            numericUpDownTo.Minimum = 1;
            numericUpDownTo.Value = compCount;
            numericUpDownFrom.Maximum = compCount - 1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Items.Add("Расчет не производился");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            Dictionary<string, double> var46;
            Dictionary<string, double> var4;
            Dictionary<string, double> var59;
            List<string>[] compData = Program.RunCalculation(Convert.ToInt32(numericUpDownFrom.Value), Convert.ToInt32(numericUpDownTo.Value),
                out var46,
                out var4,
                out var59);

            listBox1.Items.Add("Показатель 'Accounts Receivable Turnover'");
            foreach (String row in compData[0])
            {
                listBox1.Items.Add(row);
            }

            listBox1.Items.Add("Показатель 'Operating Gross Margin'");
            foreach (String row in compData[1])
            {
                listBox1.Items.Add(row);
            }
            listBox1.Items.Add("Показатель 'Cash/Current Liability'");
            foreach (String row in compData[2])
            {
                listBox1.Items.Add(row);
            }

            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            chart1.Series[2].Points.Clear();

            chart1.Series[0].Points.AddXY("Необъясненная SS", var46["dolyaSSmist"]);
            chart1.Series[0].Points.AddXY("Доля банкротов", var46["dolyaSSm"]);
            chart1.Series[0].Points.AddXY("Доля не банкротов", var46["dolyaSSw"]);

            chart1.Series[1].Points.AddXY("Необъясненная SS", var4["dolyaSSmist"]);
            chart1.Series[1].Points.AddXY("Доля банкротов", var4["dolyaSSm"]);
            chart1.Series[1].Points.AddXY("Доля не банкротов", var4["dolyaSSw"]);

            chart1.Series[2].Points.AddXY("Необъясненная SS", var59["dolyaSSmist"]);
            chart1.Series[2].Points.AddXY("Доля банкротов", var59["dolyaSSm"]);
            chart1.Series[2].Points.AddXY("Доля не банкротов", var59["dolyaSSw"]);

        }

        private void numericUpDownFrom_ValueChanged(object sender, EventArgs e)
        {
            numericUpDownTo.Minimum = numericUpDownFrom.Value + 1;
        }

        private void numericUpDownTo_ValueChanged(object sender, EventArgs e)
        {
            numericUpDownTo.Maximum = numericUpDownFrom.Value + 1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
               

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oSheet.Name = "Результаты";

                // creating Dialog
                SaveFileDialog saveFileDialog = new SaveFileDialog()
                {
                    Filter = "Excel Files | *.xlsx",
                    DefaultExt = "xlsx",
                    Title = "Save to Excel",
                    AddExtension = true // dialog automatically adds extention to file name
                };

                DialogResult result = saveFileDialog.ShowDialog();

                if (result != DialogResult.OK && result != DialogResult.Yes)
                    return;

                oSheet.Cells[1, 1] = "Для Диапазонов";
                oSheet.Cells[2, 1] = "От:";
                oSheet.Cells[2, 2] = numericUpDownFrom.Value;
                oSheet.Cells[3, 1] = "До:";
                oSheet.Cells[3, 2] = numericUpDownTo.Value;          

                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    string[] firstArray = listBox1.GetItemText(listBox1.Items[i]).Split(new char[]{':', '\t' });
                    for (int j = 0; j < firstArray.Length; j++) {
                        oSheet.Cells[i + 6, j + 1] = Convert.ToString(firstArray[j]).Replace(',', '.');
                    }
      
                }
                CreateChart(oSheet, "E6:F9", "Диаграмма по столбцу 47");
                CreateChart(oSheet, "E13:F15", "Диаграмма по столбцу 6");
                CreateChart(oSheet, "E19:F21", "Диаграмма по столбцу 60");
                oXL.Visible = true;

                oSheet.Columns.AutoFit();
                oXL.Visible = false;
                oXL.DisplayAlerts = false;
                oXL.UserControl = true;
                oWB.SaveCopyAs(saveFileDialog.FileName);
                oWB.Close();
                oXL.Quit();
                oXL.DisplayAlerts = true;
                MessageBox.Show("Отчет сохранен!");
            }
            catch (Exception theException)
            {
                String errorMessage = "Error: ";

                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
            }

        }

        private void CreateChart(Excel._Worksheet oWS, string chart_range, string name_page)
        {
            Excel._Workbook oWB;
            Excel.Range oResizeRange;
            Excel._Chart oChart;

            //Add a Chart for the selected data. 
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            oResizeRange = oWS.get_Range(chart_range);
            oChart.SetSourceData(oResizeRange);
            oChart.ChartType = Excel.XlChartType.xlColumnClustered;
            
            oChart.Location(Excel.XlChartLocation.xlLocationAsNewSheet, name_page);

        }


    }

}
