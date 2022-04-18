﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MathNet.Numerics.Distributions;

namespace OptionPricerWorkBook
{
    internal class ImpliedVolatility
    {
        private double mkt_optionPrice { get; set; }
        private double strike { get; set; }
        private double days { get; set; }
        private double spotPrice { get; set; }
        private double riskFree_rate { get; set; }
        private double tol { get; set; }
        private double divi_yield { get; set; }

        public ImpliedVolatility(double mkt_op_price, double strike_p, double total_days, double spot, double rf_rate, double err, double yield)
        {
            mkt_optionPrice = mkt_op_price;
            strike = strike_p;
            days = total_days;
            spotPrice = spot;
            riskFree_rate = rf_rate;
            tol = err;
            divi_yield = yield;

        }

        public double newton_vol(int psi)
        {
            double implied_vol = 0.1; //Suppose an initial 10% implied volatility.
            double ratio = (tol + 1); //so the while loop condition evaluates to true and condition passed.
            double days_in_year = 365;

            double price_err = 0;


            while (Math.Abs(ratio) > tol)
            {
                double d_1 = ((Math.Log(spotPrice / strike) + (riskFree_rate - divi_yield + 0.5 * Math.Pow(implied_vol, 2)) * days / days_in_year) / (implied_vol * Math.Sqrt(days / days_in_year)));
                double d_2 = (d_1 - implied_vol * Math.Sqrt(days / days_in_year));
                double vega = (spotPrice * Math.Pow(Math.E, -divi_yield * days / days_in_year) * Math.Sqrt(days / days_in_year) * Normal.PDF(0, 1, d_1));
                double N_d1 = Normal.CDF(0, 1, psi * d_1);
                double N_d2 = Normal.CDF(0, 1, psi * d_2);
                double theo_price = psi * (spotPrice * Math.Pow(Math.E, -divi_yield * days / days_in_year) * N_d1 - strike * Math.Pow(Math.E, -riskFree_rate * days / days_in_year) * N_d2);
                price_err = theo_price - mkt_optionPrice;

                ratio = price_err / vega;  //divide the price error by vega to increase the convergence rate.
                implied_vol -= ratio;
            }

            return Math.Round(implied_vol * 100, 5);
        }
    }
}
