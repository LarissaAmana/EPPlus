/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class CompileResult
    {
        private static CompileResult _empty = new CompileResult(null, DataType.Empty);

        public static CompileResult Empty
        {
            get { return _empty; }
        }

		private double? _ResultNumeric;

        public CompileResult(object result, DataType dataType, ParsingScope c = null)
        {
            Result = result;
            DataType = dataType;
            _context = c;
        }

        public CompileResult(eErrorType errorType)
        {
            Result = ExcelErrorValue.Create(errorType);
            DataType = DataType.ExcelError;
        }

        public CompileResult(ExcelErrorValue errorValue)
        {
            Require.Argument(errorValue).IsNotNull("errorValue");
            Result = errorValue;
            DataType = DataType.ExcelError;
        }

        public object Result
        {
            get;
            private set;
        }

        private ParsingScope _context;

        public object ResultValue
        {
            get
            {
                var r = Result as ExcelDataProvider.IRangeInfo;
                if (r == null)
                {
                    return Result;
                }
                else
                {
                    return r.GetValue(r.Address._fromRow, r.Address._fromCol);
                }
            }
        }

        public double ResultNumeric
        {
            get
            {
				// We assume that Result does not change unless it is a range.
				if (_ResultNumeric == null)
				{
					if (IsNumeric)
					{
						_ResultNumeric = Result == null ? 0 : Convert.ToDouble(Result);
					}
					else if (Result is DateTime datetime)
					{
						_ResultNumeric = datetime.ToOADate();
					}
					else if (Result is TimeSpan timespan)
					{
						_ResultNumeric = DateTime.FromOADate(0).Add(timespan).ToOADate();
					}
					else if (Result is ExcelDataProvider.IRangeInfo range)
					{
					    
                        if (CheckRangeCalculationPreConditions(range))
                        {
                            return GetValueForCurrentContextCell(range);
                        }

					    var c = ((ExcelDataProvider.IRangeInfo)Result).FirstOrDefault();

                        if (c == null)
						{
							return 0;
						}
						else
						{
							return c.ValueDoubleLogical;
						}
					}
				    else if (DataType == DataType.ExcelError)
				        return double.NaN;

				    // The IsNumericString and IsDateString properties will set _ResultNumeric for efficiency so we just need
				    // to check them here.
				    else if (!IsNumericString && !IsDateString)
				    {
				        _ResultNumeric = 0;
				    }
				}
				return _ResultNumeric.Value;
            }
        }

        public DataType DataType
        {
            get;
            private set;
        }
        
        public bool IsNumeric
        {
            get 
            {
                return DataType == DataType.Decimal || DataType == DataType.Integer || DataType == DataType.Empty || DataType == DataType.Boolean || DataType == DataType.Date; 
            }
        }

        public bool IsNumericString
        {
            get
            {
				double result;
				if (DataType == DataType.String && ConvertUtil.TryParseNumericString(Result, out result))
				{
					_ResultNumeric = result;
					return true;
				}
				return false;
            }
        }

		public bool IsDateString
		{
			get
			{
				DateTime result;
				if (DataType == DataType.String && ConvertUtil.TryParseDateString(Result, out result))
				{
					_ResultNumeric = result.ToOADate();
					return true;
				}
				return false;
			}
		}

		public bool IsResultOfSubtotal { get; set; }

        public bool IsHiddenCell { get; set; }

        private bool CheckRangeCalculationPreConditions(ExcelDataProvider.IRangeInfo range)
        {
            if (_context != null)
            {
                return (range.Address._fromRow != range.Address._toRow && range.Address._fromCol == range.Address._toCol
                        || range.Address._fromCol != range.Address._toCol && range.Address._fromRow == range.Address._toRow);

            }

            return false;
        }



        private double GetValueForCurrentContextCell(ExcelDataProvider.IRangeInfo range)
        {
            var value = range.Address._fromRow == range.Address._toRow
                ? range.GetValue(range.Address._fromRow, _context.Address.FromCol)
                : range.GetValue(_context.Address.FromRow, range.Address._fromCol);

            DataType = GetDataTypeForValue(value);

            var newCompileResult = new CompileResult(value, DataType, _context);

            return newCompileResult.ResultNumeric;
        }


        private DataType GetDataTypeForValue(object value)
        {

            if (value == null)
                return DataType.Empty;

            return IsNumber(value)
                ? DataType.Decimal
                : DataType.ExcelError;

            
            
        }

        private static bool IsNumber(object value)
            => value is sbyte
               || value is byte
               || value is short
               || value is ushort
               || value is int
               || value is uint
               || value is long
               || value is ulong
               || value is float
               || value is double
               || value is decimal;

    }
}
