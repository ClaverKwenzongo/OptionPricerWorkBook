using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace OptionPricerWorkBook
{
    internal class riskMetrics
    {
        static DateTime getDate(string _date)
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            DateTime wsDate = DateTime.Parse(_date, cultureInfo);

            return wsDate;
        }

        static double tenor(string start, string end)
        {
            CultureInfo culture = new CultureInfo("es-ES");

            DateTime end_date = DateTime.Parse(end, culture);
            DateTime start_date = DateTime.Parse(start, culture);

            TimeSpan interval = end_date - start_date;
            double days = (double)interval.TotalDays;

            return days;
        }

        static int getCol_Increase()
        {
            int col_count = 1;
            int col3 = 2;
            while (string.IsNullOrWhiteSpace(Globals.Sheet2.Cells[2, col3].Value?.ToString()) == false)
            {
                col_count++;
                col3++;
            }

            return col_count;
        }

        public double getPandL(string mat_date, int share_col,int date_row, string shareName, int psi, double strike)
        {
            getImpliedVol impl_vol = new getImpliedVol();
            getSharePrice s0 = new getSharePrice();
            getRates rf = new getRates();
            getYield div_yield = new getYield();

            int col_inc = getCol_Increase();

            double p_l = 0;
            int row = date_row;
            //for(int i = date_row; i < date_row+1; i++)
            //{
            //if(string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[row+1,share_col].Value?.ToString())==false)
            //{
            DateTime date_start_1 = getDate(Globals.Sheet1.Cells[row, 1].Value.ToString());
            DateTime date_start_2 = getDate(Globals.Sheet1.Cells[row+1, 1].Value.ToString());


            //string date_start_1 = Globals.Sheet1.Cells[row, 1].Value.ToString();
            //string date_start_2 = Globals.Sheet1.Cells[row+1, 1].Value.ToString();
            double Tenor_1 = tenor(mat_date, date_start_1.ToString("dd/MM/yyyy"));
            double Tenor_2 = tenor(mat_date, date_start_2.ToString("dd/MM/yyyy"));

            EuropeanOptionPricer pricer_1 = new EuropeanOptionPricer(strike, psi, Tenor_1);
            EuropeanOptionPricer pricer_2 = new EuropeanOptionPricer(strike, psi, Tenor_2);

            double impl_vol_1 = impl_vol.getImpl_Vol(Tenor_1, date_start_1.ToString("dd/MM/yyyy"), col_inc, shareName.ToUpper());
            double impl_vol_2 = impl_vol.getImpl_Vol(Tenor_2, date_start_2.ToString("dd/MM/yyyy"), col_inc, shareName.ToUpper());

            double div_yield_1 = div_yield.getDiv_Yield(Tenor_1, date_start_1.ToString("dd/MM/yyyy"), col_inc, shareName.ToUpper());
            double div_yield_2 = div_yield.getDiv_Yield(Tenor_2, date_start_2.ToString("dd/MM/yyyy"), col_inc, shareName.ToUpper());

            double rate_1 = rf.getRate(Tenor_1, date_start_1.ToString("dd/MM/yyyy"));
            double rate_2 = rf.getRate(Tenor_2, date_start_2.ToString("dd/MM/yyyy"));

            double spot_1 = s0._getSharePrice(date_start_1.ToString("dd/MM/yyyy"), shareName.ToUpper());
            double spot_2 = s0._getSharePrice(date_start_2.ToString("dd/MM/yyyy"), shareName.ToUpper());



            double price_1 = pricer_1.optionPrice(spot_1, rate_1, impl_vol_1, div_yield_1);
            double price_2 = pricer_2.optionPrice(spot_2, rate_2, impl_vol_2, div_yield_2);

            p_l = Math.Log(price_2 / price_1);
           // }
                //row++;
            //}

            return p_l;
        }

        public double getMonteCarloVaR()
        {
            double MCVaR = 0;

            return MCVaR;
        }
    }
}
