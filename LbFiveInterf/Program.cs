using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LbFiveInterf
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        public static int GetCompanyCount()
        {
            List<Company> inputData = DBConnector.ReadDataFromDB();
            return inputData.Count;
        }

        public static List<string>[] RunCalculation(int companyFrom, int companyTo, 
            out Dictionary<string, double> var46,
            out Dictionary<string, double> var4,
            out Dictionary<string, double> var59)
        {
            /*    string workFile = "C:\\#\\InputData.xlsm";
                List<Company> inputData = ReadDataExcelFileDOM(workFile);

             //   DBConnector.WriteDataDB(inputData);
            */

            List<Company> inputData = DBConnector.ReadDataFromDB();

            int RowNum,
                curCompIndex = -1,
                WNum = 0, MNum = 0;
            var46 = new Dictionary<string, double>();
            var4 = new Dictionary<string, double>();
            var59 = new Dictionary<string, double>();

            var46.Add("TotalM", 0);
            var46.Add("TotalW", 0);

            var4.Add("TotalM", 0);
            var4.Add("TotalW", 0);

            var59.Add("TotalM", 0);
            var59.Add("TotalW", 0);

            foreach (Company company in inputData)
            {
                curCompIndex++;
                if (curCompIndex < companyFrom || curCompIndex > companyTo)
                    continue;

                if (company.getBankrupt())
                {
                    var46["TotalM"] += company.getAccountsReceivableTurnover();
                    var4["TotalM"] += company.getOperatingGrossMargin();
                    var59["TotalM"] += company.getCashCurrentLiability();
                    MNum++;
                }
                else
                {
                    var46["TotalW"] += company.getAccountsReceivableTurnover();
                    var4["TotalW"] += company.getOperatingGrossMargin();
                    var59["TotalW"] += company.getCashCurrentLiability();
                    WNum++;
                }
            }
            RowNum = WNum + MNum;

            countValuesOne(var46, MNum, WNum, RowNum);
            countValuesOne(var4, MNum, WNum, RowNum);
            countValuesOne(var59, MNum, WNum, RowNum);

            curCompIndex = -1;

            foreach (Company company in inputData)
            {
                curCompIndex++;
                if (curCompIndex < companyFrom || curCompIndex > companyTo)
                    continue;

                if (company.getBankrupt())
                {   //SSm += (cell.value - AverM)^2
                    var46["SSm"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["AverM"]), 2);
                    var4["SSm"] += Math.Pow((company.getOperatingGrossMargin() - var4["AverM"]), 2);
                    var59["SSm"] += Math.Pow((company.getCashCurrentLiability() - var59["AverM"]), 2);

                    //SS += (cell.value - Average)^2
                    var46["SS"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["Average"]), 2);
                    var4["SS"] += Math.Pow((company.getOperatingGrossMargin() - var4["Average"]), 2);
                    var59["SS"] += Math.Pow((company.getCashCurrentLiability() - var59["Average"]), 2);
                }
                else
                {
                    //SSw += (cell.value - AverW)^2
                    var46["SSw"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["AverW"]), 2);
                    var4["SSw"] += Math.Pow((company.getOperatingGrossMargin() - var4["AverW"]), 2);
                    var59["SSw"] += Math.Pow((company.getCashCurrentLiability() - var59["AverW"]), 2);

                    //SS += (cell.value - Average)^2
                    var46["SS"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["Average"]), 2);
                    var4["SS"] += Math.Pow((company.getOperatingGrossMargin() - var4["Average"]), 2);
                    var59["SS"] += Math.Pow((company.getCashCurrentLiability() - var59["Average"]), 2);
                }
            }

            countValuesTwo(var46, MNum, WNum);
            countValuesTwo(var4, MNum, WNum);
            countValuesTwo(var59, MNum, WNum);

            List<string> accRecTList = WriteData(var46);
            List<string> opeGrMgnList = WriteData(var4);
            List<string> cshLbltList = WriteData(var59);

            return new List<string>[] { accRecTList, opeGrMgnList, cshLbltList };
        }


        private static List<string> WriteData(Dictionary<string, double> variable)
        {
            List<string> outputStr = new List<string>();

            outputStr.Add("Влияние показателя на выходную переменную: " + Convert.ToString(Math.Round(variable["D"], 2)) + "%" +
                "\t\t\tНеобъясненная SS: " + Convert.ToString(Math.Round(variable["dolyaSSmist"], 2)));
            outputStr.Add("Общая сумма квадратов отклонений: " + Convert.ToString(Math.Round(variable["SS"], 2)) +
                "\t\t\tДоля банкротов в общей ошибке: " + Convert.ToString(Math.Round(variable["dolyaSSm"], 2)));
            outputStr.Add("Объясненная влиянием 'а' сум.кв.откл: " + Convert.ToString(Math.Round(variable["SSeff"], 2)) +
                "\t\t\tДоля не банкротов в общей ошибке: " + Convert.ToString(Math.Round(variable["dolyaSSw"], 2)));
            outputStr.Add("Необъясненная сумма квадратов отклонений: " + Convert.ToString(Math.Round(variable["SSmist"], 2)));
            outputStr.Add("\n");

            return outputStr;
        }
        private static void countValuesOne(Dictionary<string, double> variable, int MNum, int WNum, int RowNum)
        {

            variable.Add("AverM", variable["TotalM"] / MNum);                    //AverM = TotalM / MNum
            variable.Add("AverW", variable["TotalW"] / WNum);                   //AverW = TotalW / WNum
            variable.Add("Total", variable["TotalW"] + variable["TotalM"]);   //Total = TotalW + TotalM
            variable.Add("Average", variable["Total"] / RowNum);                //Average = Total / RowNum

            variable.Add("SSw", 0);
            variable.Add("SSm", 0);
            variable.Add("SS", 0);
        }

        private static void countValuesTwo(Dictionary<string, double> variable, int MNum, int WNum)
        {
            variable.Add("SSmist", variable["SSw"] + variable["SSm"]);                             //SSmist = SSw + SSm
            variable.Add("SSeff", WNum * Math.Pow(variable["AverW"] - variable["Average"], 2) +        //SSeff = WNum * (AverW - Average) ^ 2 + _
                                            MNum * Math.Pow(variable["AverM"] - variable["Average"], 2));     //   MNum * (AverM - Average) ^ 2
            variable.Add("D", variable["SSeff"] / variable["SS"] * 100);                           // D = SSeff / SS * 100

            //Для графика
            variable.Add("dolyaSSmist", variable["SSmist"] / variable["SS"] * 100);
            variable.Add("dolyaSSw", variable["SSw"] / variable["SS"] * 100);
            variable.Add("dolyaSSm", variable["SSm"] / variable["SS"] * 100);

        }
    }
}
