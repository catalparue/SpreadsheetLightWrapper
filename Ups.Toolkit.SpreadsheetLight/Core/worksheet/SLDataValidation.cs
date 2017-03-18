using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    /// <summary>
    ///     Data validation types.
    /// </summary>
    public enum SLDataValidationAllowedValues
    {
        /// <summary>
        ///     Whole number.
        /// </summary>
        WholeNumber = 0,

        /// <summary>
        ///     Decimal.
        /// </summary>
        Decimal,

        /// <summary>
        ///     Date.
        /// </summary>
        Date,

        /// <summary>
        ///     Time.
        /// </summary>
        Time,

        /// <summary>
        ///     Text length.
        /// </summary>
        TextLength
    }

    /// <summary>
    ///     Data validation operations with 1 operand.
    /// </summary>
    public enum SLDataValidationSingleOperandValues
    {
        /// <summary>
        ///     Equal.
        /// </summary>
        Equal = 0,

        /// <summary>
        ///     Not equal.
        /// </summary>
        NotEqual,

        /// <summary>
        ///     Greater than.
        /// </summary>
        GreaterThan,

        /// <summary>
        ///     Less than.
        /// </summary>
        LessThan,

        /// <summary>
        ///     Greater than or equal.
        /// </summary>
        GreaterThanOrEqual,

        /// <summary>
        ///     Less than or equal.
        /// </summary>
        LessThanOrEqual
    }

    /// <summary>
    ///     Encapsulates properties and methods for data validations.
    /// </summary>
    public class SLDataValidation
    {
        internal SLDataValidation()
        {
        }

        internal bool Date1904 { get; set; }

        internal string Formula1 { get; set; }
        internal string Formula2 { get; set; }

        internal bool HasDataValidation
        {
            get
            {
                return (SequenceOfReferences.Count > 0) && ((Type != DataValidationValues.None)
                                                            || (ErrorTitle.Length > 0) || (Error.Length > 0)
                                                            || (PromptTitle.Length > 0) || (Prompt.Length > 0));
            }
        }

        internal DataValidationValues Type { get; set; }
        internal DataValidationErrorStyleValues ErrorStyle { get; set; }
        internal DataValidationImeModeValues ImeMode { get; set; }
        internal DataValidationOperatorValues Operator { get; set; }

        internal bool AllowBlank { get; set; }
        internal bool ShowDropDown { get; set; }

        /// <summary>
        ///     Specifies if the input message is shown.
        /// </summary>
        public bool ShowInputMessage { get; set; }

        /// <summary>
        ///     Specifies if the error message is shown.
        /// </summary>
        public bool ShowErrorMessage { get; set; }

        internal string ErrorTitle { get; set; }
        internal string Error { get; set; }
        internal string PromptTitle { get; set; }
        internal string Prompt { get; set; }

        internal List<SLCellPointRange> SequenceOfReferences { get; set; }

        internal void InitialiseDataValidation(int StartRowIndex, int StartColumnIndex, int EndRowIndex,
            int EndColumnIndex, bool Date1904)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            SetAllNull();
            this.Date1904 = Date1904;
            SequenceOfReferences.Add(new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex,
                iEndColumnIndex));
        }

        private void SetAllNull()
        {
            Date1904 = false;
            Formula1 = string.Empty;
            Formula2 = string.Empty;
            Type = DataValidationValues.None;
            ErrorStyle = DataValidationErrorStyleValues.Stop;
            ImeMode = DataValidationImeModeValues.NoControl;
            Operator = DataValidationOperatorValues.Between;
            AllowBlank = false;
            ShowDropDown = false;
            ShowInputMessage = true;
            ShowErrorMessage = true;
            ErrorTitle = string.Empty;
            Error = string.Empty;
            PromptTitle = string.Empty;
            Prompt = string.Empty;
            SequenceOfReferences = new List<SLCellPointRange>();
        }

        /// <summary>
        ///     Allow any value.
        /// </summary>
        public void AllowAnyValue()
        {
            Type = DataValidationValues.None;
        }

        /// <summary>
        ///     Allow only whole numbers.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(bool IsBetween, int Minimum, int Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Whole;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow only whole numbers.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(bool IsBetween, long Minimum, long Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Whole;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow only whole numbers.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Whole;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = CleanDataSourceForFormula(Minimum);
            Formula2 = CleanDataSourceForFormula(Maximum);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow only whole numbers.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(SLDataValidationSingleOperandValues DataOperator, int DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Whole;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow only whole numbers.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(SLDataValidationSingleOperandValues DataOperator, long DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Whole;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow only whole numbers.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(SLDataValidationSingleOperandValues DataOperator, string DataValue,
            bool IgnoreBlank)
        {
            Type = DataValidationValues.Whole;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = CleanDataSourceForFormula(DataValue);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, float Minimum, float Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, double Minimum, double Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, decimal Minimum, decimal Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = CleanDataSourceForFormula(Minimum);
            Formula2 = CleanDataSourceForFormula(Maximum);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, float DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, double DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, decimal DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, string DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Decimal;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = CleanDataSourceForFormula(DataValue);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow a list of values.
        /// </summary>
        /// <param name="DataSource">The data source. For example, "$A$1:$A$5"</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        /// <param name="InCellDropDown">True if a dropdown list appears for selecting. False otherwise.</param>
        public void AllowList(string DataSource, bool IgnoreBlank, bool InCellDropDown)
        {
            Type = DataValidationValues.List;
            Operator = DataValidationOperatorValues.Between;

            if (DataSource.StartsWith("="))
            {
                Formula1 = DataSource.Substring(1);
            }
            else
            {
                if (Regex.IsMatch(DataSource, "^\\s*\\$[A-Za-z]{1,3}\\$[0-9]{1,7}"))
                    Formula1 = DataSource;
                else
                    Formula1 = string.Format("\"{0}\"", DataSource);
            }

            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
            // I don't know why it's reversed. It seems to make sense when "normal"...
            ShowDropDown = !InCellDropDown;
        }

        /// <summary>
        ///     Allow date values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(bool IsBetween, DateTime Minimum, DateTime Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Date;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = SLTool.CalculateDaysFromEpoch(Minimum, Date1904).ToString(CultureInfo.InvariantCulture);
            Formula2 = SLTool.CalculateDaysFromEpoch(Maximum, Date1904).ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow date values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">
        ///     The minimum value. Any valid date formatted value is fine. It is suggested to just copy the value
        ///     you have in Excel interface.
        /// </param>
        /// <param name="Maximum">
        ///     The maximum value. Any valid date formatted value is fine. It is suggested to just copy the value
        ///     you have in Excel interface.
        /// </param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.Date;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;

            DateTime dt;

            if (Minimum.StartsWith("="))
            {
                Formula1 = Minimum.Substring(1);
            }
            else
            {
                if (DateTime.TryParse(Minimum, out dt))
                    Formula1 = SLTool.CalculateDaysFromEpoch(dt, Date1904).ToString(CultureInfo.InvariantCulture);
                else
                    Formula1 = "1";
            }

            if (Maximum.StartsWith("="))
            {
                Formula2 = Maximum.Substring(1);
            }
            else
            {
                if (DateTime.TryParse(Maximum, out dt))
                    Formula2 = SLTool.CalculateDaysFromEpoch(dt, Date1904).ToString(CultureInfo.InvariantCulture);
                else
                    Formula2 = "1";
            }

            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow date values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">
        ///     The data value. Any valid date formatted value is fine. It is suggested to just copy the value
        ///     you have in Excel interface.
        /// </param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(SLDataValidationSingleOperandValues DataOperator, DateTime DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Date;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = SLTool.CalculateDaysFromEpoch(DataValue, Date1904).ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow date values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">
        ///     The data value. Any valid date formatted value is fine. It is suggested to just copy the value
        ///     you have in Excel interface.
        /// </param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(SLDataValidationSingleOperandValues DataOperator, string DataValue, bool IgnoreBlank)
        {
            Type = DataValidationValues.Date;
            Operator = TranslateOperatorValues(DataOperator);

            DateTime dt;

            if (DataValue.StartsWith("="))
            {
                Formula1 = DataValue.Substring(1);
            }
            else
            {
                if (DateTime.TryParse(DataValue, out dt))
                    Formula1 = SLTool.CalculateDaysFromEpoch(dt, Date1904).ToString(CultureInfo.InvariantCulture);
                else
                    Formula1 = "1";
            }

            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow time values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="StartHour">The start hour between 0 to 23 (both inclusive).</param>
        /// <param name="StartMinute">The start minute between 0 to 59 (both inclusive).</param>
        /// <param name="StartSecond">The start second between 0 to 59 (both inclusive).</param>
        /// <param name="EndHour">The end hour between 0 to 23 (both inclusive).</param>
        /// <param name="EndMinute">The end minute between 0 to 59 (both inclusive).</param>
        /// <param name="EndSecond">The end second between 0 to 59 (both inclusive).</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(bool IsBetween, int StartHour, int StartMinute, int StartSecond, int EndHour,
            int EndMinute, int EndSecond, bool IgnoreBlank)
        {
            if (StartHour < 0) StartHour = 0;
            if (StartHour > 23) StartHour = 23;
            if (StartMinute < 0) StartMinute = 0;
            if (StartMinute > 59) StartMinute = 59;
            if (StartSecond < 0) StartSecond = 0;
            if (StartSecond > 59) StartSecond = 59;
            if (EndHour < 0) EndHour = 0;
            if (EndHour > 23) EndHour = 23;
            if (EndMinute < 0) EndMinute = 0;
            if (EndMinute > 59) EndMinute = 59;
            if (EndSecond < 0) EndSecond = 0;
            if (EndSecond > 59) EndSecond = 59;

            Type = DataValidationValues.Time;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;

            double fTime = 0;

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            fTime = StartHour/24.0 + StartMinute/1440.0 + StartSecond/86400.0;
            Formula1 = fTime.ToString(CultureInfo.InvariantCulture);

            fTime = EndHour/24.0 + EndMinute/1440.0 + EndSecond/86400.0;
            Formula2 = fTime.ToString(CultureInfo.InvariantCulture);

            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow time values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="StartTime">
        ///     The start time. Any valid time formatted value is fine. It is suggested to just copy the value
        ///     you have in Excel interface.
        /// </param>
        /// <param name="EndTime">
        ///     The end time. Any valid time formatted value is fine. It is suggested to just copy the value you
        ///     have in Excel interface.
        /// </param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(bool IsBetween, string StartTime, string EndTime, bool IgnoreBlank)
        {
            Type = DataValidationValues.Time;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;

            double fTime = 0;
            DateTime dt;
            string sTime;
            // we include the day, month and year for formatting because it seems that parsing based
            // just on time (hour, minute, second, AM/PM designator) is too much for TryParseExact()...
            string[] saFormats =
            {
                "dd/MM/yyyy H", "dd/MM/yyyy h t", "dd/MM/yyyy h tt", "dd/MM/yyyy H:m",
                "dd/MM/yyyy h:m t", "dd/MM/yyyy h:m tt", "dd/MM/yyyy H:m:s", "dd/MM/yyyy h:m:s t", "dd/MM/yyyy h:m:s tt"
            };
            var sSampleDate = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            if (StartTime.StartsWith("="))
            {
                Formula1 = StartTime.Substring(1);
            }
            else
            {
                sTime = string.Format("{0} {1}", sSampleDate, StartTime);
                if (DateTime.TryParseExact(sTime, saFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    fTime = dt.Hour/24.0 + dt.Minute/1440.0 + dt.Second/86400.0;
                    Formula1 = fTime.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    Formula1 = "0";
                }
            }

            if (EndTime.StartsWith("="))
            {
                Formula2 = EndTime.Substring(1);
            }
            else
            {
                sTime = string.Format("{0} {1}", sSampleDate, EndTime);
                if (DateTime.TryParseExact(sTime, saFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    fTime = dt.Hour/24.0 + dt.Minute/1440.0 + dt.Second/86400.0;
                    Formula1 = fTime.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    Formula1 = "0";
                }
            }

            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow time values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Hour">The hour between 0 to 23 (both inclusive).</param>
        /// <param name="Minute">The minute between 0 to 59 (both inclusive).</param>
        /// <param name="Second">The second between 0 to 59 (both inclusive).</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(SLDataValidationSingleOperandValues DataOperator, int Hour, int Minute, int Second,
            bool IgnoreBlank)
        {
            if (Hour < 0) Hour = 0;
            if (Hour > 23) Hour = 23;
            if (Minute < 0) Minute = 0;
            if (Minute > 59) Minute = 59;
            if (Second < 0) Second = 0;
            if (Second > 59) Second = 59;

            Type = DataValidationValues.Time;
            Operator = TranslateOperatorValues(DataOperator);

            double fTime = 0;

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            fTime = Hour/24.0 + Minute/1440.0 + Second/86400.0;
            Formula1 = fTime.ToString(CultureInfo.InvariantCulture);

            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow time values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Time">
        ///     The time. Any valid time formatted value is fine. It is suggested to just copy the value you have in
        ///     Excel interface.
        /// </param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(SLDataValidationSingleOperandValues DataOperator, string Time, bool IgnoreBlank)
        {
            Type = DataValidationValues.Time;
            Operator = TranslateOperatorValues(DataOperator);

            double fTime = 0;
            DateTime dt;
            string sTime;
            // we include the day, month and year for formatting because it seems that parsing based
            // just on time (hour, minute, second, AM/PM designator) is too much for TryParseExact()...
            string[] saFormats =
            {
                "dd/MM/yyyy H", "dd/MM/yyyy h t", "dd/MM/yyyy h tt", "dd/MM/yyyy H:m",
                "dd/MM/yyyy h:m t", "dd/MM/yyyy h:m tt", "dd/MM/yyyy H:m:s", "dd/MM/yyyy h:m:s t", "dd/MM/yyyy h:m:s tt"
            };
            var sSampleDate = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            if (Time.StartsWith("="))
            {
                Formula1 = Time.Substring(1);
            }
            else
            {
                sTime = string.Format("{0} {1}", sSampleDate, Time);
                if (DateTime.TryParseExact(sTime, saFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    fTime = dt.Hour/24.0 + dt.Minute/1440.0 + dt.Second/86400.0;
                    Formula1 = fTime.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    Formula1 = "0";
                }
            }

            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow data according to text length.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(bool IsBetween, int Minimum, int Maximum, bool IgnoreBlank)
        {
            if (Minimum < 0) Minimum = 0;
            if (Maximum < 0) Maximum = 0;

            Type = DataValidationValues.TextLength;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow data according to text length.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            Type = DataValidationValues.TextLength;
            Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            Formula1 = CleanDataSourceForFormula(Minimum);
            Formula2 = CleanDataSourceForFormula(Maximum);
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow data according to text length.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Length">The text length for comparison.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(SLDataValidationSingleOperandValues DataOperator, int Length, bool IgnoreBlank)
        {
            if (Length < 0) Length = 0;

            Type = DataValidationValues.TextLength;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = Length.ToString(CultureInfo.InvariantCulture);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow data according to text length.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Length">The text length for comparison.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(SLDataValidationSingleOperandValues DataOperator, string Length, bool IgnoreBlank)
        {
            Type = DataValidationValues.TextLength;
            Operator = TranslateOperatorValues(DataOperator);
            Formula1 = CleanDataSourceForFormula(Length);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Allow custom validation.
        /// </summary>
        /// <param name="Formula">The formula used for validation.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowCustom(string Formula, bool IgnoreBlank)
        {
            Type = DataValidationValues.Custom;
            Operator = DataValidationOperatorValues.Between;
            Formula1 = CleanDataSourceForFormula(Formula);
            Formula2 = string.Empty;
            AllowBlank = IgnoreBlank;
        }

        /// <summary>
        ///     Set the input message.
        /// </summary>
        /// <param name="Title">The title of the input message.</param>
        /// <param name="Message">The input message.</param>
        public void SetInputMessage(string Title, string Message)
        {
            PromptTitle = Title;
            Prompt = Message;
        }

        /// <summary>
        ///     Set the error alert.
        /// </summary>
        /// <param name="ErrorStyle">The error style.</param>
        /// <param name="Title">The title of the error alert.</param>
        /// <param name="Message">The error message.</param>
        public void SetErrorAlert(DataValidationErrorStyleValues ErrorStyle, string Title, string Message)
        {
            this.ErrorStyle = ErrorStyle;
            ErrorTitle = Title;
            Error = Message;
        }

        /// <summary>
        ///     Set the error alert.
        /// </summary>
        /// <param name="Title">The title of the error alert.</param>
        /// <param name="Message">The error message.</param>
        public void SetErrorAlert(string Title, string Message)
        {
            ErrorStyle = DataValidationErrorStyleValues.Stop;
            ErrorTitle = Title;
            Error = Message;
        }

        internal string CleanDataSourceForFormula(string DataValue)
        {
            var result = DataValue;
            if (result.StartsWith("="))
                result = DataValue.Substring(1);

            return result;
        }

        internal DataValidationOperatorValues TranslateOperatorValues(SLDataValidationSingleOperandValues Operator)
        {
            var result = DataValidationOperatorValues.Between;
            switch (Operator)
            {
                case SLDataValidationSingleOperandValues.Equal:
                    result = DataValidationOperatorValues.Equal;
                    break;
                case SLDataValidationSingleOperandValues.NotEqual:
                    result = DataValidationOperatorValues.NotEqual;
                    break;
                case SLDataValidationSingleOperandValues.GreaterThan:
                    result = DataValidationOperatorValues.GreaterThan;
                    break;
                case SLDataValidationSingleOperandValues.LessThan:
                    result = DataValidationOperatorValues.LessThan;
                    break;
                case SLDataValidationSingleOperandValues.GreaterThanOrEqual:
                    result = DataValidationOperatorValues.GreaterThanOrEqual;
                    break;
                case SLDataValidationSingleOperandValues.LessThanOrEqual:
                    result = DataValidationOperatorValues.LessThanOrEqual;
                    break;
            }

            return result;
        }

        internal void FromDataValidation(DataValidation dv)
        {
            SetAllNull();

            if (dv.Formula1 != null) Formula1 = dv.Formula1.Text;
            if (dv.Formula2 != null) Formula2 = dv.Formula2.Text;

            if (dv.Type != null) Type = dv.Type.Value;
            if (dv.ErrorStyle != null) ErrorStyle = dv.ErrorStyle.Value;
            if (dv.ImeMode != null) ImeMode = dv.ImeMode.Value;
            if (dv.Operator != null) Operator = dv.Operator.Value;
            if (dv.AllowBlank != null) AllowBlank = dv.AllowBlank.Value;
            if (dv.ShowDropDown != null) ShowDropDown = dv.ShowDropDown.Value;
            if (dv.ShowInputMessage != null) ShowInputMessage = dv.ShowInputMessage.Value;
            if (dv.ShowErrorMessage != null) ShowErrorMessage = dv.ShowErrorMessage.Value;

            if (dv.ErrorTitle != null) ErrorTitle = dv.ErrorTitle.Value;
            if (dv.Error != null) Error = dv.Error.Value;
            if (dv.PromptTitle != null) PromptTitle = dv.PromptTitle.Value;
            if (dv.Prompt != null) Prompt = dv.Prompt.Value;

            // it has to be not-null because it's a required thing, but you never know...
            if (dv.SequenceOfReferences != null)
                SequenceOfReferences = SLTool.TranslateSeqRefToCellPointRange(dv.SequenceOfReferences);
        }

        internal DataValidation ToDataValidation()
        {
            var dv = new DataValidation();

            if (Formula1.Length > 0) dv.Formula1 = new Formula1(Formula1);
            if (Formula2.Length > 0) dv.Formula2 = new Formula2(Formula2);

            if (Type != DataValidationValues.None) dv.Type = Type;
            if (ErrorStyle != DataValidationErrorStyleValues.Stop) dv.ErrorStyle = ErrorStyle;
            if (ImeMode != DataValidationImeModeValues.NoControl) dv.ImeMode = ImeMode;
            if (Operator != DataValidationOperatorValues.Between) dv.Operator = Operator;

            if (AllowBlank) dv.AllowBlank = AllowBlank;
            if (ShowDropDown) dv.ShowDropDown = ShowDropDown;
            if (ShowInputMessage) dv.ShowInputMessage = ShowInputMessage;
            if (ShowErrorMessage) dv.ShowErrorMessage = ShowErrorMessage;

            if (ErrorTitle.Length > 0) dv.ErrorTitle = ErrorTitle;
            if (Error.Length > 0) dv.Error = Error;

            if (PromptTitle.Length > 0) dv.PromptTitle = PromptTitle;
            if (Prompt.Length > 0) dv.Prompt = Prompt;

            dv.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(SequenceOfReferences);

            return dv;
        }

        internal SLDataValidation Clone()
        {
            var dv = new SLDataValidation();
            dv.Date1904 = Date1904;
            dv.Formula1 = Formula1;
            dv.Formula2 = Formula2;
            dv.Type = Type;
            dv.ErrorStyle = ErrorStyle;
            dv.ImeMode = ImeMode;
            dv.Operator = Operator;
            dv.AllowBlank = AllowBlank;
            dv.ShowDropDown = ShowDropDown;
            dv.ShowInputMessage = ShowInputMessage;
            dv.ShowErrorMessage = ShowErrorMessage;
            dv.ErrorTitle = ErrorTitle;
            dv.Error = Error;
            dv.PromptTitle = PromptTitle;
            dv.Prompt = Prompt;

            dv.SequenceOfReferences = new List<SLCellPointRange>();
            foreach (var pt in SequenceOfReferences)
                dv.SequenceOfReferences.Add(pt);

            return dv;
        }
    }
}