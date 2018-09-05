using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Roman : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var argumentList = arguments.ToList();
            ValidateArguments(argumentList, 1, 2, eErrorType.Name);
            int arabic = checkAndGetFirstArgument(arguments);
            string result = String.Empty;

            if (arabic == 0) //Roman(0)
            {
                return new CompileResult(result, DataType.String);
            }
            
            if (argumentList.Count == 2)
            {
                int intArg2;

                //if (Object.ReferenceEquals(argumentList[1].Value.GetType(), bolArg2.GetType()))
                if (argumentList[1].Value is bool boolArgument)
                {
                    result = ToRomanNumber(arabic, boolArgument ? 0 : 4);
                }
                else
                {
                    intArg2 = ArgToInt(argumentList, 1);
                    result = ToRomanNumber(arabic, intArg2);
                }
            }
            else 
            {
                result = ToRomanNumber(arabic, 0);
            }



            return CreateResult(result, DataType.String);
        }

        private int checkAndGetFirstArgument(IEnumerable<FunctionArgument> arguments)
        {
            double arabicdouble = 0;
            try
            {
                arabicdouble = ArgToDecimal(arguments, 0);
            }
            catch (InvalidCastException)
            {
                throw new Exceptions.ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Name));
            }
         

            int arabic = 0;
            if (!(Int32.TryParse((System.Math.Truncate(arabicdouble).ToString()), out arabic)))
            {
                throw new Exceptions.ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
            }
            if (arabic > 3999)
            {
                throw new Exceptions.ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
            }
            return arabic;
        }



        private static string ToRomanNumber(int arabic, int param)
        {

            if (param > 4)
            {
                throw new Exceptions.ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
            }
            else if (param == 0)
            {
                return ToRomanNumber(arabic, _divChainParam0, new Dictionary<int, string>());
            }
            else if (param == 1)
            {

                return ToRomanNumber(arabic, _divChainParam1, _dicToRoman1);
            }
            else if (param == 2)
            {

                return ToRomanNumber(arabic, _divChainParam2, _dicToRoman2);
            }
            else if (param == 3)
            {

                return ToRomanNumber(arabic, _divChainParam3, _dicToRoman3);
            }
            else if (param == 4)
            {

                return ToRomanNumber(arabic, _divChainParam4, _dicToRoman4);
            }


            return null;

        }




        private static string ToRomanNumber(int arabic, int[] divChain, Dictionary<int, string> dicToRoman)
        {
            int[] digits = new int[divChain.Length];


            for (int i = 0; i < digits.Length; i++)
            {
                digits[i] = 0;
            }

            for (int i = 0; i < divChain.Length; i++)
            {
                digits[i] = arabic / divChain[i];
                arabic = arabic % divChain[i];
            }

            StringBuilder returnValue = new StringBuilder();

            for (int i = 0; i < digits.Length; i++)
            {
                string append = "";

                if ((divChain[i] / 1000) >= 1)
                {
                    if (digits[i] != 0)
                        _dicToRomanClassic1000.TryGetValue(divChain[i], out append);
                }
                else if (dicToRoman.Count > 0 && dicToRoman.ContainsKey(divChain[i]))
                {
                    if (digits[i] == 0)
                    {
                        continue;
                    }

                    dicToRoman.TryGetValue(divChain[i], out append);


                }
                else if ((divChain[i] / 100) >= 1)
                {
                    if (digits[i] == 0)
                        continue;
                    _dicToRomanClassic100.TryGetValue(divChain[i], out append);
                }
                else if ((divChain[i] / 10) >= 1)
                {
                    if (digits[i] == 0)
                        continue;
                    _dicToRomanClassic10.TryGetValue(divChain[i], out append);
                }
                else if ((divChain[i] / 1) >= 1)
                {
                    if (digits[i] == 0)
                        continue;
                    _dicToRomanClassic1.TryGetValue(divChain[i], out append);
                }
                else
                {
                    throw new Exceptions.ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
                }

                returnValue.Append(append);
            }




            return returnValue.ToString();

        }


        private static Dictionary<int, string> _dicToRomanClassic1 = new Dictionary<int, string>()
        {
            {0, "" },
            { 1, "I"},
            { 2, "II"},
            { 3, "III"},
            { 4, "IV"},
            { 5, "V"},
            { 9, "IX"},
        };



        private static Dictionary<int, string> _dicToRomanClassic10 = new Dictionary<int, string>()
        {
            {0, "" },
            { 10, "X"},
            { 20, "XX"},
            { 30, "XXX"},
            { 40, "XL"},
            { 50, "L"},
            { 90, "XC"},
        };

        private static Dictionary<int, string> _dicToRomanClassic100 = new Dictionary<int, string>()
        {
            {0, "" },
            { 100, "C"},
            { 200, "CC"},
            { 300, "CCC"},
            { 400, "CD"},
            { 500, "D"},
            { 900, "CM"},
        };

        private static Dictionary<int, string> _dicToRomanClassic1000 = new Dictionary<int, string>()
        {
            {0, "" },
            { 1000, "M"},
            { 2000, "MM"},
            { 3000, "MMM"},
        };

        private static Dictionary<int, string> _dicToRoman1 = new Dictionary<int, string>()
        {

            {0, "" }, //0 => should never happen
            { 950, "LM"}, //950
            { 450, "LD"}, //450
            { 95, "VC"}, //95
            { 45, "VL"}, //45
        };

        private static Dictionary<int, string> _dicToRoman2 = new Dictionary<int, string>()
        {

            { 0, "" }, //0 => should never happen
            { 990, "XM"}, //990
            { 950, "LM"}, //950
            { 490, "XD"}, //490
            { 450, "LD"}, //450
            { 99, "IC"}, //99
            { 95, "VC"}, //95
            { 49, "IL"}, //49
            { 45, "VL"}, //45
        };

        private static Dictionary<int, string> _dicToRoman3 = new Dictionary<int, string>()
            {

                { 0, "" }, //0 => should never happen
                { 995, "VM"}, //995
                { 990, "XM"}, //990
                { 950, "LM"}, //950
                { 495, "VD"}, //495
                { 490, "XD"}, //490
                { 450, "LD"}, //450
                { 99, "IC"}, //99
                { 95, "VC"}, //95
                { 49, "IL"}, //49
                { 45, "VL"}, //45
            };


        private static Dictionary<int, string> _dicToRoman4 = new Dictionary<int, string>()
        {

            { 0, "" }, //0 => should never happen
            { 999, "IM"}, //999
            { 995, "VM"}, //995
            { 990, "XM"}, //990
            { 950, "LM"}, //950
            { 499, "ID"}, //499
            { 495, "VD"}, //495
            { 490, "XD"}, //490
            { 450, "LD"}, //450
            { 99, "IC"}, //99
            { 95, "VC"}, //95
            { 49, "IL"}, //49
            { 45, "VL"}, //45
        };

        private static int[] _divChainParam0 = { 3000, 2000, 1000, 900, 500, 400, 300, 200, 100, 90, 50, 40, 30, 20, 10, 9, 5, 4, 3, 2, 1 };
        private static int[] _divChainParam1 = { 3000, 2000, 1000, 950, 900, 500, 450, 400, 300, 200, 100, 95, 90, 50, 45, 40, 30, 20, 10, 9, 5, 4, 3, 2, 1 };
        private static int[] _divChainParam2 = { 3000, 2000, 1000, 990, 950, 900, 500, 490, 450, 400, 300, 200, 100, 99, 95, 90, 50, 49, 45, 40, 30, 20, 10, 9, 5, 4, 3, 2, 1 };
        private static int[] _divChainParam3 = { 3000, 2000, 1000, 995, 990, 950, 900, 500, 495, 490, 450, 400, 300, 200, 100, 99, 95, 90, 50, 49, 45, 40, 30, 20, 10, 9, 5, 4, 3, 2, 1 };
        private static int[] _divChainParam4 = { 3000, 2000, 1000, 999, 995, 990, 950, 900, 500, 499, 495, 490, 450, 400, 300, 200, 100, 99, 95, 90, 50, 49, 45, 40, 30, 20, 10, 9, 5, 4, 3, 2, 1 };

    }
}
