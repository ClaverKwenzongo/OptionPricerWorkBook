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
            Globals.Sheet4.Columns.ColumnWidth = 20;
            Globals.Sheet4.Cells[row_start, 3].Value = "Portfolio Value";
            Globals.Sheet4.Cells[row_start+2, 3].Value = "Share Name";
            Globals.Sheet4.Cells[row_start+3, 3].Value = "Start Date";
            Globals.Sheet4.Cells[row_start + 4, 3].Value = "Maturity date";
            Globals.Sheet4.Cells[row_start + 5, 3].Value = "Strike Price";
            Globals.Sheet4.Cells[row_start+6, 3].Value = "Position (short or long)";
            Globals.Sheet4.Cells[row_start+7, 3].Value = "Option Type (put or call)";
            Globals.Sheet4.Cells[row_start+8, 3].Value = "Amount";

            Globals.Sheet4.Cells[row_start + 11, 3].Value = "Option Price";
            Globals.Sheet4.Cells[row_start + 12, 3].Value = "Delta";
            Globals.Sheet4.Cells[row_start + 13, 3].Value = "Gamma";
            Globals.Sheet4.Cells[row_start + 14, 3].Value = "Vega";
            Globals.Sheet4.Cells[row_start + 15, 3].Value = "Implied Volatility";

            Globals.Sheet4.Cells[row_start + 20, 3].Value = "Portfolio Risk Metrics";
            Globals.Sheet4.Cells[row_start + 21, 2].Value = "Historical VaR";
            Globals.Sheet4.Cells[row_start + 22, 3].Value = "1% VaR";
            Globals.Sheet4.Cells[row_start + 23, 3].Value = "2% VaR";
            Globals.Sheet4.Cells[row_start + 24, 2].Value = "Monte Carlo VaR";
            Globals.Sheet4.Cells[row_start + 25, 3].Value = "1% VaR";
            Globals.Sheet4.Cells[row_start + 26, 3].Value = "2% VaR";






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
            getRates getRates = new getRates(); 

            int col_inc = getCol_Increase();

            double portfolio_val = 0;

            int j = 4;
            while(string.IsNullOrWhiteSpace(Globals.Sheet4.Cells[row_start+2,j].Value?.ToString()) == false)
            {
                string myStartDate = Globals.Sheet4.Cells[row_start + 3, j].Value.ToString();
                Debug.WriteLine(myStartDate);
                string user_share = Globals.Sheet4.Cells[row_start + 2, j].Value;
                double S_0 = getShare._getSharePrice(myStartDate, user_share.ToUpper());

                double Tenor = tenor(Globals.Sheet4.Cells[row_start + 3, j].Value.ToString(), Globals.Sheet4.Cells[row_start + 4, j].Value.ToString());

                double impl_Vol = getVol.getImpl_Vol(Tenor, myStartDate, col_inc, user_share.ToUpper());

                double div_yield = getDividentYield.getDiv_Yield(Tenor, myStartDate, col_inc, user_share.ToUpper());

                double rate = getRates.getRate(Tenor, myStartDate);

                double K = Globals.Sheet4.Cells[row_start + 5, j].Value;

                string option_type = Globals.Sheet4.Cells[row_start + 7, j].Value;

                int psi = 0;
                if (option_type.ToUpper() == "PUT")
                {
                    psi = -1;
                }
                else
                {
                    psi = 1;
                }

                string option_pos = Globals.Sheet4.Cells[row_start + 6, j].Value;
                double size = Globals.Sheet4.Cells[row_start + 8, j].Value;

                if(option_pos.ToUpper() == "SHORT")
                {
                    size = -size;
                }
                else
                {
                    size = size;
                }


                EuropeanOptionPricer pricer = new EuropeanOptionPricer(K, psi, Tenor);

                double mkt_op_price = pricer.optionPrice(S_0, rate, impl_Vol, div_yield);
                ImpliedVolatility newton_implied_vol = new ImpliedVolatility(mkt_op_price, K, Tenor, S_0, rate, Math.Pow(10, -8), div_yield);

                Globals.Sheet4.Cells[row_start + 11, j].Value = pricer.optionPrice(S_0, rate, impl_Vol, div_yield);
                Globals.Sheet4.Cells[row_start + 12, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Delta");
                Globals.Sheet4.Cells[row_start + 13, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Gamma");
                Globals.Sheet4.Cells[row_start + 14, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Vega");
                Globals.Sheet4.Cells[row_start + 15, j].Value = newton_implied_vol.newton_vol(psi);

                portfolio_val += size * pricer.optionPrice(S_0, rate, impl_Vol, div_yield);

                j++;
            }

            var val = string.Format("{0:C}", portfolio_val);

            Globals.Sheet4.Cells[row_start, 4].Value = val;
        }
    }
}
