namespace SpreadsheetLightWrapper.Core.conditionalformatting
{
    /// <summary>
    ///     Conditional formatting options for five icon sets.
    /// </summary>
    public class SLFiveIconSetOptions
    {
        internal bool IsCustomIcon;

        internal SLIconValues vIcon1;

        internal SLIconValues vIcon2;

        internal SLIconValues vIcon3;

        internal SLIconValues vIcon4;

        internal SLIconValues vIcon5;

        /// <summary>
        ///     Initializes an instance of SLFiveIconSetOptions.
        /// </summary>
        /// <param name="IconSetType">The type of icon set.</param>
        public SLFiveIconSetOptions(SLFiveIconSetValues IconSetType)
        {
            this.IconSetType = IconSetType;
            ReverseIconOrder = false;
            ShowIconOnly = false;

            IsCustomIcon = false;

            GreaterThanOrEqual2 = true;
            GreaterThanOrEqual3 = true;
            GreaterThanOrEqual4 = true;
            GreaterThanOrEqual5 = true;

            switch (IconSetType)
            {
                case SLFiveIconSetValues.FiveArrows:
                    vIcon1 = SLIconValues.RedDownArrow;
                    vIcon2 = SLIconValues.YellowDownInclineArrow;
                    vIcon3 = SLIconValues.YellowSideArrow;
                    vIcon4 = SLIconValues.YellowUpInclineArrow;
                    vIcon5 = SLIconValues.GreenUpArrow;
                    break;
                case SLFiveIconSetValues.FiveArrowsGray:
                    vIcon1 = SLIconValues.GrayDownArrow;
                    vIcon2 = SLIconValues.GrayDownInclineArrow;
                    vIcon3 = SLIconValues.GraySideArrow;
                    vIcon4 = SLIconValues.GrayUpInclineArrow;
                    vIcon5 = SLIconValues.GrayUpArrow;
                    break;
                case SLFiveIconSetValues.FiveBoxes:
                    vIcon1 = SLIconValues.ZeroFilledBoxes;
                    vIcon2 = SLIconValues.OneFilledBox;
                    vIcon3 = SLIconValues.TwoFilledBoxes;
                    vIcon4 = SLIconValues.ThreeFilledBoxes;
                    vIcon5 = SLIconValues.FourFilledBoxes;
                    break;
                case SLFiveIconSetValues.FiveQuarters:
                    vIcon1 = SLIconValues.WhiteCircleAllWhiteQuarters;
                    vIcon2 = SLIconValues.CircleWithThreeWhiteQuarters;
                    vIcon3 = SLIconValues.CircleWithTwoWhiteQuarters;
                    vIcon4 = SLIconValues.CircleWithOneWhiteQuarter;
                    vIcon5 = SLIconValues.BlackCircle;
                    break;
                case SLFiveIconSetValues.FiveRating:
                    vIcon1 = SLIconValues.SignalMeterWithNoFilledBars;
                    vIcon2 = SLIconValues.SignalMeterWithOneFilledBar;
                    vIcon3 = SLIconValues.SignalMeterWithTwoFilledBars;
                    vIcon4 = SLIconValues.SignalMeterWithThreeFilledBars;
                    vIcon5 = SLIconValues.SignalMeterWithFourFilledBars;
                    break;
            }

            Value2 = "20";
            Value3 = "40";
            Value4 = "60";
            Value5 = "80";

            Type2 = SLConditionalFormatRangeValues.Percent;
            Type3 = SLConditionalFormatRangeValues.Percent;
            Type4 = SLConditionalFormatRangeValues.Percent;
            Type5 = SLConditionalFormatRangeValues.Percent;
        }

        internal SLFiveIconSetValues IconSetType { get; set; }

        /// <summary>
        ///     Specifies if the icons in the set are reversed.
        /// </summary>
        public bool ReverseIconOrder { get; set; }

        /// <summary>
        ///     Specifies if only the icon is shown. Set to false to show both icon and value.
        /// </summary>
        public bool ShowIconOnly { get; set; }

        /// <summary>
        ///     The 1st icon.
        /// </summary>
        public SLIconValues Icon1
        {
            get { return vIcon1; }
            set
            {
                if (vIcon1 != value)
                {
                    vIcon1 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        ///     The 2nd icon.
        /// </summary>
        public SLIconValues Icon2
        {
            get { return vIcon2; }
            set
            {
                if (vIcon2 != value)
                {
                    vIcon2 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        ///     The 3rd icon.
        /// </summary>
        public SLIconValues Icon3
        {
            get { return vIcon3; }
            set
            {
                if (vIcon3 != value)
                {
                    vIcon3 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        ///     The 4th icon.
        /// </summary>
        public SLIconValues Icon4
        {
            get { return vIcon4; }
            set
            {
                if (vIcon4 != value)
                {
                    vIcon4 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        ///     The 5th icon.
        /// </summary>
        public SLIconValues Icon5
        {
            get { return vIcon5; }
            set
            {
                if (vIcon5 != value)
                {
                    vIcon5 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        ///     Specifies if values are to be greater than or equal to the 2nd range value. Set to false if values are to be
        ///     strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual2 { get; set; }

        /// <summary>
        ///     Specifies if values are to be greater than or equal to the 3rd range value. Set to false if values are to be
        ///     strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual3 { get; set; }

        /// <summary>
        ///     Specifies if values are to be greater than or equal to the 4th range value. Set to false if values are to be
        ///     strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual4 { get; set; }

        /// <summary>
        ///     Specifies if values are to be greater than or equal to the 5th range value. Set to false if values are to be
        ///     strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual5 { get; set; }

        /// <summary>
        ///     The 2nd range value.
        /// </summary>
        public string Value2 { get; set; }

        /// <summary>
        ///     The 3rd range value.
        /// </summary>
        public string Value3 { get; set; }

        /// <summary>
        ///     The 4th range value.
        /// </summary>
        public string Value4 { get; set; }

        /// <summary>
        ///     The 5th range value.
        /// </summary>
        public string Value5 { get; set; }

        /// <summary>
        ///     The conditional format type for the 2nd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type2 { get; set; }

        /// <summary>
        ///     The conditional format type for the 3rd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type3 { get; set; }

        /// <summary>
        ///     The conditional format type for the 4th range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type4 { get; set; }

        /// <summary>
        ///     The conditional format type for the 5th range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type5 { get; set; }
    }
}