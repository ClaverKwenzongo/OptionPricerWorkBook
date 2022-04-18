using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Globalization;
using System.Diagnostics;

namespace OptionPricerWorkBook
{
    public partial class Ribbon1
    {
        int row_start = 5;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.Sheet4.Cells[row_start+2, 3].Value = "Share Name";
            Globals.Sheet4.Cells[row_start+3, 3].Value = "Start Date";
            Globals.Sheet4.Cells[row_start + 4, 3].Value = "Maturity date";
            Globals.Sheet4.Cells[row_start+5, 3].Value = "Position";
            Globals.Sheet4.Cells[row_start+6, 3].Value = "Option Type";
            Globals.Sheet4.Cells[row_start+7, 3].Value = "Amount";


        }

        static int getCol_Increase()
        {
            int col_count = 1;
            int col3 = 2;
            while (string.IsNullOrWhiteSpace(Globals.Sheet2.Cells[2,col3].Value?.ToString()) == false)
            {
                col_count++;
                col3++;
            }

            return col_count;
        }

        static double tenor( string start, string end)
        {
            CultureInfo culture = new CultureInfo("es-ES");

            DateTime end_date = DateTime.Parse(end, culture);
            DateTime start_date = DateTime.Parse(start, culture);

            TimeSpan interval = end_date - start_date;
            double days = (double)interval.TotalDays;

            return days;
        }

            private void testBtn_Click(object sender, RibbonControlEventArgs e)
        {
            getSharePrice getShare = new getSharePrice();
            getImpliedVol getVol = new getImpliedVol();
            getYield getDividentYield = new getYield();

            string myStartDate = Globals.Sheet4.Cells[row_start+3,4].Value.ToString();
            Debug.WriteLine(myStartDate);
            string user_share = Globals.Sheet4.Cells[row_start+2,4].Value;
            double s0 = getShare._getSharePrice(myStartDate, user_share.ToUpper());

            //Globals.Sheet4.Cells[3,4].Value = 
            int col_inc = getCol_Increase();
            double Tenor = tenor(Globals.Sheet4.Cells[row_start+3,4].Value.ToString(),Globals.Sheet4.Cells[row_start+4,4].Value.ToString());
            double impl_Vol = getVol.getImpl_Vol(Tenor, myStartDate, col_inc, user_share.ToUpper());

            double div_yield = getDividentYield.getDiv_Yield(Tenor, myStartDate, col_inc, user_share.ToUpper());

            //Globals.Sheet4.Cells[3, 4].Value = div_yield;

            //using (ExcelWorksheet ws = Globals.Sheet2 as ExcelWorksheet)
            //{
            //}
            //int div_start_col = getStartCol(Globals.Sheet3, user_share.ToUpper());
        }
    }
}
