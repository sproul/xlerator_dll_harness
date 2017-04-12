using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;
using System.Data;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using XLeratorDLL_financial;

namespace Monr
{
    using wct = XLeratorDLL_financial.XLeratorDLL_financial;
    using FinancialTypes = XLeratorDLL_financial.FinancialTypes;

    class Program
    {
        const string vbCrLf = "\r\n";
        const char vbTab = '\t';
        static bool only_correct = false;
        static bool spectrum_mode = false;
        static bool terse_mode = false;
        static string invocation_stem = "xlerator.yieldn ";
        static string invocation_suffix = "";
        static List<int> Bases = new List<int>();
        static List<int> Frequencies = new List<int>();
        static List<int> Frequencies_general = new List<int> { 1, 2, 4, 12 };
        static List<int> Frequencies_364 = new List<int> { 7, 14, 28, 91, 182, 364 };
        static Dictionary<string, Double> Parms_to_yield = new Dictionary<string, Double>();
        static Dictionary<string, Double> Parms_to_yield_poor_matches = new Dictionary<string, Double>(); // store poor results here if running in only_correct mode
        static Dictionary<int, string> Frequency_code_to_doc = new Dictionary<int, string>()
                {
                    { 1, "annual" },
                    { 2, "semi-annual" },
                    { 4, "quarterly" },
                    //Frequency_code_to_doc_general.Add(6, "bi-monthly") // documented but leads to a throw
                    { 12, "monthly" },
                    
                    
                    
                    
                    { 364, "every 52 wx" },
                    { 182, "every 26 wx" },
                    { 91, "every 13 wx" },
                    { 28, "every 4 wx" },
                    { 14, "every 2 wx" },
                    { 7, "weekly" }
                };
        static Dictionary<int, string> Basis_code_to_doc = new Dictionary<int, string>()
            {
                    { 0, "BOND US (NASD) 30/360" },
                    { 1, "ACTUAL Actual/Actual" },
                    { 2, "A360 Actual/360" },
                    { 3, "A365 Actual/365" },
                    { 4, "30E/360 (ISDA), 30E/360, ISDA, 30E/360 ISDA, EBOND  European 30/360" },
                    { 5, "30/360, 30/360 ISDA, GERMAN 30/360 ISDA" },
                    { 6, "NL/ACT No Leap Year/ACT" },
                    { 7, "NL/365 No Leap Year /365" },
                    { 8, "NL/360 No Leap Year /360" },
                    { 9, "A/365 Actual/364" },
                    { 10, "BOND NON-EOM US (NASD) 30/360 non-end-of-month" },
                    { 11, "ACTUAL NON-EOM Actual/Actual non-end-of-month" },
                    { 12, "A360 NON-EOM Actual/360 non-end-of-month" },
                    { 13, "A365 NON-EOM Actual/365 non-end-of-month" },
                    { 14, "30E/360 NON-EOM, 30E/360 ICMA NONEOM, EBOND NON-EOM, European 30/360 non-end-of-month" },
                    { 15, "30/360 NON-EOM, 30/360 ISDA NONEOM, GERMAN NON-EOM, 30/360 ISDA non-end-of-month" },
                    { 16, "NL/ACT NON-EOM No Leap Year/ACT non-end-of-month" },
                    { 17, "NL/365 NON-EOM No Leap Year/365 non-end-of-month" },
                    { 18, "NL/360 NON-EOM No Leap Year/360 non-end-of-month" },
                    { 19, "A/365 NON-EOM Actual/364 non-end-of-month" }
            };
        private static List<int> All_basis_codes = new List<int>() { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19 };

        static void Main(string[] args)
        {
            // Invocation must be structured with the following ordering:
            // 1.) Command flags
            // 2.) Price or price/target_yield pairs
            // 3.) settlement maturity rate redemption [first_coupon_date last_coupon_date issue_date]
            int x = 0;
            string arg;
            bool flag_args_done = false;
            while (!flag_args_done && x < args.Length)
            {
                arg = args[x++];
                switch (arg)
                {
                    case "-b":
                    case "-basis":
                        Program.Bases.Add(Parse_basis(args[x++]));
                        break;
                    case "-f":
                    case "-frequency":
                        Program.Frequencies.Add(Parse_frequency(args[x++]));
                        break;
                    case "-only_correct":
                        Program.only_correct = true;
                        break;
                    case "-ref":
                        // this program does nothing w/ this info, but it affords an opportunity for the web UI to note what EMMA page is the basis for slurped
                        // form fields, if any (or record any other interesting reference data, which would then be visible at the top of the output)
                        x++;    // actual arg to be ignored
                        break;
                    case "-spectrum":
                        Program.Spectrum_test();
                        return;
                    case "-suppress_deviations_gt":
                        throw new Exception("IMPL");
                    case "-t":
                        Program.terse_mode = true;
                        break;
                    case "-terse":
                        Program.terse_mode = true;
                        break;
                    default:
                        flag_args_done = true;
                        x--;    // we didn't parse this arg, so keep it available
                        break;
                }
            }
            Add_to_stem(args, 0, x);
            if (x >= args.Length)
            {
                throw new Exception("not enough args (" + args.Length + ")");
            }
            arg = args[x];
            Dictionary<Double, Double> Price_to_target_yield = new Dictionary<Double, Double>();
            if (!arg.Contains("/"))
            {
                Price_to_target_yield.Add(Parse_double(arg), 0);
                x++;
            }
            else
            {
                do
                {
                    Double Price = Parse_double(new Regex("/.*").Replace(arg, ""));
                    Double Target_yield = Parse_double(new Regex(".*/").Replace(arg, ""));
                    Price_to_target_yield.Add(Price, Target_yield);
                    arg = args[++x];
                } while (arg.Contains("/"));
            }
            if (args.Length - x < 8)
            {
                Line("bad number of args (" + args.Length + ", " + x + " consumed)");
                return;
            }
            Add_to_suffix(args, x, args.Length - x);
            int[] Settlement_date_vals = new int[3] { Parse_int(args[x++]), Parse_int(args[x++]), Parse_int(args[x++]) };
            int[] Maturity_date_vals = new int[3] { Parse_int(args[x++]), Parse_int(args[x++]), Parse_int(args[x++]) };
            Double Rate = Parse_double(args[x++]);
            Double Redemption = Parse_double(args[x++]);
            // optional args follow
            int[] FirstCoupon_date_vals = null;
            int[] LastCoupon_date_vals = null;
            int[] Issue_date_vals = null;
            if (args.Length - x > 2)
            {
                Issue_date_vals = new int[3] { Parse_int(args[x++]), Parse_int(args[x++]), Parse_int(args[x++]) };
            }
            if (args.Length - x > 2)
            {
                FirstCoupon_date_vals = new int[3] { Parse_int(args[x++]), Parse_int(args[x++]), Parse_int(args[x++]) };
            }
            if (args.Length - x > 2)
            {
                LastCoupon_date_vals = new int[3] { Parse_int(args[x++]), Parse_int(args[x++]), Parse_int(args[x++]) };
            }
            foreach (var Price_to_target_yield_item in Price_to_target_yield)
            {
                Double Price = Price_to_target_yield_item.Key;
                Double Target_yield = Price_to_target_yield_item.Value;

                Explore_yields(Settlement_date_vals,
                           Maturity_date_vals,
                           Rate,
                           Price,
                           Redemption,
                           Target_yield,
                           Issue_date_vals,
                           FirstCoupon_date_vals,
                           LastCoupon_date_vals);
            }
        }
    	public static void Add_to_stem(string[] args, int index, int length)
        {
            if (length - index > 0)
            {
                string[] t = new string[length - index + 1];
                Array.Copy(args, index, t, 0, length - index + 1);
                Program.invocation_stem += String.Join(" ", t) + " ";
            }
        }
        public static void Add_to_suffix(string[] args, int index, int length)
        {
            if (length - index > 0)
            {
                string[] t = new string[length - index + 1];
                Array.Copy(args, index, t, 0, length - index + 1);
                Program.invocation_suffix += " " + String.Join(" ", t);
            }
        }
        public static Dictionary<string, Double> descriptions_to_results()
        {
            if (Program.only_correct && !Program.Parms_to_yield.Any()) {
                return Program.Parms_to_yield_poor_matches;
            }
            return Program.Parms_to_yield;
        }
        public static int Parse_int(string s)
        {
            try
            {
                int i = Int32.Parse(s);
                return i;
            }
            catch (Exception e)
            {
                Line("exception (" + e + ") trying to make '" + s + "' into an int");
                throw;
            }
        }
        public static int Parse_basis(string s)
        {
            int i = Parse_int(s);
            int MIN_BASIS = 0;
            int MAX_BASIS = 19;
            if (i < MIN_BASIS || i > MAX_BASIS) {
                throw new Exception("expected '" + s + "' to be in the range " + MIN_BASIS + ".." + MAX_BASIS + " expected for basis codes");
            }
            return i;
        }
    	public static int Parse_frequency(string s)
        {
            int i = Parse_int(s);
            if (i!=1 && i!=2 && i!=4 && i!=7 && i!=12 && i!=14 && i!=28 && i!=91 && i!=182 && i!=364)
            {
                throw new Exception("expected '" + s + "' to be in the set { 1 2 4 7 12 14 28 91 182 364 } of valid frequency codes");
            }
            return i;
        }

        public static double Parse_double(string s)
        {
            try
            {
                double d = Double.Parse(s);
                return d;
            }
            catch (System.FormatException e)
            {
                string emsg = "exception (" + e + ") trying to make '" + s + "' into an double";
                Line(emsg);
                throw new Exception(emsg);
            }
        }

        public static Double Yield1(string Type,
                                int[] Settlement_date_vals,
                                int[] Maturity_date_vals,
                                Double Rate,
                                Double Price,
                                Double Redemption,
                                Double Frequency,
                                int Basis,
                                int[] Issue_date_vals = null,
                                int[] FirstCoupon_date_vals = null,
                                int[] LastCoupon_date_vals = null)
        {
            DateTime Settlement = new DateTime(Settlement_date_vals[0], Settlement_date_vals[1], Settlement_date_vals[2]);
            DateTime Maturity = new DateTime(Maturity_date_vals[0], Maturity_date_vals[1], Maturity_date_vals[2]);
            Double result;

            string Basis_string = "" + Basis;
            switch (Type)
            {
                case "clean":
                    result = wct.YIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis_string);
                    break;
                case "dirty":
                    DateTime Issue = wct.NULL_DATE;
                    DateTime FirstCoupon = wct.NULL_DATE;
                    DateTime LastCoupon = wct.NULL_DATE;

                    if (Issue_date_vals != null) Issue = new DateTime(Issue_date_vals[0], Issue_date_vals[1], Issue_date_vals[2]);
                    if (FirstCoupon_date_vals != null) FirstCoupon = new DateTime(FirstCoupon_date_vals[0], FirstCoupon_date_vals[1], FirstCoupon_date_vals[2]);
                    if (LastCoupon_date_vals != null) LastCoupon = new DateTime(LastCoupon_date_vals[0], LastCoupon_date_vals[1], LastCoupon_date_vals[2]);

                    result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis_string, Issue, FirstCoupon, LastCoupon);
                    break;
                default:
                    throw new Exception("bad type");
            }
            return result;
        }

        public static List<int> Get_basis_codes()
        {
            if (Program.Bases.Count() > 0) {
                return Program.Bases;
            }
            return Program.All_basis_codes;
        }

        public static List<int> Get_frequencies(int basis_code)
        {
            if (Program.Frequencies.Count() != 0)
            {
                return Program.Frequencies;
            }
            if (basis_code == 9) {
                return Program.Frequencies_364;
            }
            return Program.Frequencies_general;
        }
        public static void Yieldn(int[] Settlement_date_vals,
                                                        int[] Maturity_date_vals,
                                                        Double Rate,
                                                        Double Price,
                                                        double Redemption,
                                                        Double Target_yield,
                                                        int[] Issue_date_vals = null,
                                                        int[] FirstCoupon_date_vals = null,
                                                        int[] LastCoupon_date_vals = null)
        {
            string[] yield_types = Program.spectrum_mode ? new string[1] { "clean" } : new string[2] { "clean", "dirty" };
            List<int> Basis_codes = Get_basis_codes();
            foreach (var Basis in Basis_codes)
            {
                string Basis_description = Program.Basis_code_to_doc[Basis];

                List<int> Frequencies = Get_frequencies(Basis);
                foreach (var Frequency in Frequencies)
                {
                    string frequency_description = Program.Frequency_code_to_doc[Frequency];
                    foreach (var yield_type in yield_types)
                    {
                        Double result = Yield1(yield_type,
                                               Settlement_date_vals,
                                               Maturity_date_vals,
                                               Rate,
                                               Price,
                                               Redemption,
                                               Frequency,
                                               Basis,
                                               Issue_date_vals,
                                               FirstCoupon_date_vals,
                                               LastCoupon_date_vals);
                        if (!Double.IsNaN(result))
                        {
                            string parms;
                            parms = "";
                            if (!Program.spectrum_mode) {
                                parms += Rnd(result) + "\t" + yield_type + "\t";
                            }
                            parms += Frequency;
                            if (!Program.terse_mode) {
                                parms += " " + frequency_description;
                            }
                            parms += "\t" + Basis;
                            if (!Program.terse_mode) {
                                parms += " " + Basis_description;
                            }
                            if (Program.spectrum_mode) {
                                parms += "\t" + Maturity_date_vals[1] + "/" + Maturity_date_vals[2] + "/" + Maturity_date_vals[0] + "\t" + Price + "\t" + Rate;
                            }
                            // we either save the raw yield as our key result, or we save the deviation from the target_yield.
                            if (Target_yield != 0)
                            {
                                result = Math.Abs(Target_yield - result);
                                if (Program.only_correct && result > 0.005)
                                {
                                    Program.Parms_to_yield_poor_matches.Add(parms, result);
                                    continue;
                                }
                            }
                            Program.Parms_to_yield.Add(parms, result);
                        }
                    }
                }
            }
        }

        public static void Line(string s)
        {
            Console.WriteLine(s);
            Debug.WriteLine(s);
        }
        public static void Test_Explore_yields()
            {
                int[] Settlement_date_vals = new int[3] { 2017, 3, 21 };
                int[] Maturity_date_vals = new int[3] { 2020, 12, 1 };
                Double Rate = 1.59; // "coupon"
                Double Price = 108.612; // from today's trade
                Double Redemption = 100; // redemption value per face value
                int[] Issue_date_vals = new int[3] { 2017, 2, 28 };
                int[] FirstCoupon_date_vals = new int[3] { 2017, 6, 1 };
                int[] LastCoupon_date_vals = null;
                Explore_yields(Settlement_date_vals,
                               Maturity_date_vals,
                               Rate,
                               Price,
                               Redemption,
                               0,
                               Issue_date_vals,
                               FirstCoupon_date_vals,
                               LastCoupon_date_vals);
            }
            public static void Explore_yields(int[] Settlement_date_vals,
                                      int[] Maturity_date_vals,
                                      Double Rate,
                                      Double Price,
                                      Double Redemption,
                                      Double Target_yield,
                                      int[] Issue_date_vals,
                                      int[] FirstCoupon_date_vals,
                                      int[] LastCoupon_date_vals)
            {
                Line(Program.invocation_stem + Price + "/" + Target_yield + Program.invocation_suffix);
                Yieldn(Settlement_date_vals, Maturity_date_vals, Rate, Price, Redemption, Target_yield, Issue_date_vals, FirstCoupon_date_vals, LastCoupon_date_vals);
                int lineCount = 0;
                var result_items = Program.descriptions_to_results().ToList();
                result_items.Sort((pair1, pair2) => pair1.Value.CompareTo(pair2.Value));
                if (Program.only_correct) {
                    result_items = PruneResults(result_items);
                }

                string header = "line\t" + (Target_yield != 0 ? "error\t" : "") + "rate\ttype\tfrequency\tbasis";
                Line(header);

                foreach (var result_item in result_items)
                {
                    lineCount++;
                    string description = result_item.Key;
                    if (Target_yield == 0)
                    {
                        // note that the yield is included in the description, so no need to add it to the line
                        //Double deviation_from_target_yield = result_item.Value;



                        Line(lineCount + "\t" + description);
                    }
                    else
                    {
                        Double yield = result_item.Value;
                        Line(lineCount + "\t" + Rnd(yield) + "\t" + description);
                    }
                }
            }

        	private static List<DateTime> Create_maturity_dates(DateTime Settlement)
            {
                List<DateTime> maturity_dates = new List<DateTime>();
                maturity_dates.Add(Settlement.AddDays(7));
                maturity_dates.Add(Settlement.AddDays(15));
                maturity_dates.Add(Settlement.AddMonths(1));
                maturity_dates.Add(Settlement.AddMonths(3));
                maturity_dates.Add(Settlement.AddMonths(6));
                maturity_dates.Add(Settlement.AddYears(1));
                maturity_dates.Add(Settlement.AddYears(3));
                maturity_dates.Add(Settlement.AddYears(5));
                maturity_dates.Add(Settlement.AddYears(10));
                return maturity_dates;
            }
            public static void Spectrum_test()
            {
                Program.spectrum_mode = true;
                Program.terse_mode = true;
                int[] Settlement_date_vals = { 2016, 3, 23 };
                DateTime Settlement = new DateTime(Settlement_date_vals[0], Settlement_date_vals[1], Settlement_date_vals[2]);
                List<DateTime> Maturity_dates = Create_maturity_dates(Settlement);
                Double[] Rates = new Double[] { 1.2, 2.2, 3.2, 4.2, 5.2, 6.2, 8.2, 10.2 };
                Double[] Prices = new Double[] { 90, 95, 100, 105, 110, 115, 120 };
                Double Redemption = 100;
                foreach (var Maturity_date in Maturity_dates) 
                {
                    foreach (var Rate in Rates) 
                    {
                        foreach (var Price in Prices) 
                        {
                            int[] Maturity_date_vals = new int[3] { Maturity_date.Year, Maturity_date.Month,  Maturity_date.Day };
                            Yieldn(Settlement_date_vals, Maturity_date_vals, Rate, Price, Redemption, 0, null, null, null );
                            break;
                        }
                        break;
                    }
                    break;
                }
                int lineCount = 0;
                
                foreach (var result_item in Program.descriptions_to_results().ToList())
                {
                    lineCount++;
                    string description = result_item.Key;
                    Double yield = result_item.Value;
                    Line(lineCount + "\t" + Rnd(yield) + "\t" + description);
                }
            }

        private static List<KeyValuePair<string, double>> PruneResults(List<KeyValuePair<string, double>> fullList)
        {
            List<KeyValuePair<string, double>> prunedList = new List<KeyValuePair<string, double>>();
            const int TARGET_COUNT = 5;
            // in the event of a tie for the lead, don't prune away co-winners
            double winning_value = fullList[0].Value;
            do
            {
                prunedList.Add(fullList[0]);
                fullList.RemoveAt(0);
            } while (fullList.Count() > 0 && fullList[0].Value == winning_value);
            while (prunedList.Count() < TARGET_COUNT && fullList.Count() > 0)
            {
                prunedList.Add(fullList[0]);
                fullList.RemoveAt(0);
            }
            return prunedList;
        }

        public static string Rnd(Double d)
            {
                return "" + Math.Round(d, 2);
            }
	public static void Test_DIRTYYIELD()
            {
                Double Rate;
                Double Price;
                double Frequency;
                double Redemption;
                int Basis;
                Double result;

                int[] Settlement = new int[3] { 2017, 3, 21 };
                int[] Maturity = new int[3] { 2020, 12, 1 };
                Rate = 1.59; // "coupon"
                Price = 108.612; // from today's trade
                Redemption = 100; // redemption value per face value
                Frequency = 4; // payments per year
                Basis = 1; // daycount convention, from table
                int[] Issue = new int[3] { 2017, 2, 28 };
                int[] FirstCoupon = new int[3] { 2017, 6, 1 };
                int[] LastCoupon = null;

                result = Yield1("clean", Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis);
 
                Debug.WriteLine("YIELD manual : RPI (expect 1.59): " + result.ToString());
                result = Yield1("dirty", Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);
                
                Debug.WriteLine("DIRTYYIELD manual : RPI (expect 1.59): " + result.ToString());
                return;
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                

            /*
            //=====================
            // EXAMPLE# 1
            //=====================

            //Regular coupons
            Settlement = new DateTime(2014, 5, 1);
            Maturity = new DateTime(2034, 6, 15);
            Rate = 0.025;
            Price = 96.9453139716364;
            Redemption = 100;
            Frequency = 2;
            Basis = "1";
            Issue = wct.NULL_DATE;
            FirstCoupon = wct.NULL_DATE;
            LastCoupon = wct.NULL_DATE;

            result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);

            Console.WriteLine("DIRTYYIELD : RPI : " + result.ToString());   //96.9453139716364
            File.WriteAllText("DIRTYYIELD.csv", "RPI : " + result.ToString());

            //=====================
            // EXAMPLE# 2
            //=====================

            //Odd Short First
            Settlement = new DateTime(2014, 5, 1);
            Maturity = new DateTime(2034, 6, 15);
            Rate = 0.025;
            Price = 96.0075631077824;
            Redemption = 100;
            Frequency = 2;
            Basis = "1";
            Issue = new DateTime(2014, 5, 1);
            FirstCoupon = new DateTime(2014, 6, 15);

            result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);

            Console.WriteLine("DIRTYYIELD : Odd Short First : " + result.ToString());   //96.0075631077824
            File.AppendAllText("DIRTYYIELD.csv", vbCrLf + "Odd Short First : " + result.ToString());

            //=====================
            // EXAMPLE# 3
            //=====================

            //Odd Long First
            Settlement = new DateTime(2014, 5, 1);
            Maturity = new DateTime(2034, 6, 15);
            Rate = 0.025;
            Price = 96.0033702877755;
            Redemption = 100;
            Frequency = 2;
            Basis = "1";
            Issue = new DateTime(2014, 5, 1);
            FirstCoupon = new DateTime(2014, 12, 15);

            result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);

            Console.WriteLine("DIRTYYIELD : Odd Long First : " + result.ToString());   //96.0033702877755
            File.AppendAllText("DIRTYYIELD.csv", vbCrLf + "Odd Long First : " + result.ToString());

            //=====================
            // EXAMPLE# 4
            //=====================

            //Odd Short Last
            Settlement = new DateTime(2014, 10, 1);
            Maturity = new DateTime(2014, 12, 15);
            Rate = 0.0225;
            Price = 100.544776634675;
            Redemption = 100;
            Frequency = 2;
            Basis = "1";
            Issue = wct.NULL_DATE;
            FirstCoupon = wct.NULL_DATE;
            LastCoupon = new DateTime(2014, 9, 15);

            result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);

            Console.WriteLine("DIRTYYIELD : Odd Short Last : " + result.ToString());   //100.544776634675
            File.AppendAllText("DIRTYYIELD.csv", vbCrLf + "Odd Short Last : " + result.ToString());

            //=====================
            // EXAMPLE# 5
            //=====================

            //Odd Long Last
            Settlement = new DateTime(2014, 10, 1);
            Maturity = new DateTime(2014, 12, 15);
            Rate = 0.0225;
            Price = 101.669543602845;
            Redemption = 100;
            Frequency = 2;
            Basis = "1";
            Issue = wct.NULL_DATE;
            FirstCoupon = wct.NULL_DATE;
            LastCoupon = new DateTime(2014, 3, 15);

            result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);

            Console.WriteLine("DIRTYYIELD : Odd Long Last : " + result.ToString());    //101.669543602845
            File.AppendAllText("DIRTYYIELD.csv", vbCrLf + "Odd Long Last : " + result.ToString());

            //=====================
            // EXAMPLE# 6
            //=====================

            //Odd long first & odd long last
            Settlement = new DateTime(2013, 3, 4);
            Maturity = new DateTime(2022, 11, 28);
            Rate = 0.03125;
            Price = 104.420159644839;
            Redemption = 100;
            Frequency = 2;
            Basis = "1";
            Issue = new DateTime(2012, 6, 7);
            FirstCoupon = new DateTime(2013, 3, 15);
            LastCoupon = new DateTime(2022, 3, 15);

            result = wct.DIRTYYIELD(Settlement, Maturity, Rate, Price, Redemption, Frequency, Basis, Issue, FirstCoupon, LastCoupon);

            Console.WriteLine("DIRTYYIELD : Odd Long First & Last : " + result.ToString());
            File.AppendAllText("DIRTYYIELD.csv", vbCrLf + "Odd Long First & Last : " + result.ToString());
            */
        }
        static void Bondamort()
        {
            DateTime Settlement;
            DateTime Maturity;
            Double Rate;
            Double FaceAmount;
            Double CleanPrice;
            Double Redemption;
            Double Frequency;
            String Basis;
            DateTime IssueDate;
            DateTime FirstInterestDate;
            DateTime LastInterestDate;
            FinancialTypes.BONDAMORT_table result;
            StringBuilder retStr;

            //=====================
            // EXAMPLE# 1
            //=====================

            // We buy 1,000,000 in face value of a bond on 2012-05-03 at a price of 999000.
            // The bond matures on 2012-06-30 and accrues interest using the Actual/365 day-count method.
            // The interest rate is 5%.

            Settlement = new DateTime(2012, 5, 3);
            Maturity = new DateTime(2012, 6, 30);
            Rate = 0.05;
            FaceAmount = 1000000;
            CleanPrice = 999000;
            Redemption = wct.NULL_DOUBLE;
            Frequency = 2;
            Basis = "3";
            IssueDate = wct.NULL_DATE;
            FirstInterestDate = wct.NULL_DATE;
            LastInterestDate = wct.NULL_DATE;

            result = wct.BONDAMORT(Settlement, Maturity, Rate, FaceAmount, CleanPrice, Redemption, Frequency, Basis, IssueDate, FirstInterestDate, LastInterestDate);

            // Write the table output to a csv
            retStr = new StringBuilder();
            foreach (DataColumn col in result.Columns)
            {
                retStr.Append(col.ColumnName + ",");
            }
            retStr.Remove(retStr.Length - 1, 1);
            retStr.AppendLine();
            foreach (DataRow row in result.Rows)
            {
                for (int i = 0; i < result.Columns.Count; i++)
                {
                    retStr.Append(row[i].ToString());
                    if (i < result.Columns.Count - 1)
                    {
                        retStr.Append(",");
                    }
                    else
                    {
                        retStr.AppendLine();
                    }
                }
            }
            Console.WriteLine("BONDAMORT : ex1 : " + retStr.ToString());

            if (System.Diagnostics.Debugger.IsAttached)
            {
                // Keep the console window open in debug mode.
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }
    }
}
