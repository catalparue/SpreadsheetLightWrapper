namespace SpreadsheetLightWrapper.Core.conditionalformatting
{
    /// <summary>
    ///     Conditional formatting options for four icon sets.
    /// </summary>
    public class SLFourIconSetOptions
    {
        internal bool IsCustomIcon;

        internal SLIconValues vIcon1;

        internal SLIconValues vIcon2;

        internal SLIconValues vIcon3;

        internal SLIconValues vIcon4;

        /// <summary>
        ///     Initializes an instance of SLFourIconSetOptions.
        /// </summary>
        /// <param name="IconSetType">The type of icon set.</param>
        public SLFourIconSetOptions(SLFourIconSetValues IconSetType)
        {
            this.IconSetType = IconSetType;
            ReverseIconOrder = false;
            ShowIconOnly = false;

            IsCustomIcon = false;

            GreaterThanOrEqual2 = true;
            GreaterThanOrEqual3 = true;
            GreaterThanOrEqual4 = true;

            switch (IconSetType)
            {
                case SLFourIconSetValues.FourArrows:
                    vIcon1 = SLIconValues.RedDownArrow;
                    vIcon2 = SLIconValues.YellowDownInclineArrow;
                    vIcon3 = SLIconValues.YellowUpInclineArrow;
                    vIcon4 = SLIconValues.GreenUpArrow;
                    break;
                case SLFourIconSetValues.FourArrowsGray:
                    vIcon1 = SLIconValues.GrayDownArrow;
                    vIcon2 = SLIconValues.GrayDownInclineArrow;
                    vIcon3 = SLIconValues.GrayUpInclineArrow;
                    vIcon4 = SLIconValues.GrayUpArrow;
                    break;
                case SLFourIconSetValues.FourRating:
                    vIcon1 = SLIconValues.SignalMeterWithOneFilledBar;
                    vIcon2 = SLIconValues.SignalMeterWithTwoFilledBars;
                    vIcon3 = SLIconValues.SignalMeterWithThreeFilledBars;
                    vIcon4 = SLIconValues.SignalMeterWithFourFilledBars;
                    break;
                case SLFourIconSetValues.FourRedToBlack:
                    vIcon1 = SLIconValues.BlackCircle;
                    vIcon2 = SLIconValues.GrayCircle;
                    vIcon3 = SLIconValues.PinkCircle;
                    vIcon4 = SLIconValues.RedCircle;
                    break;
                case SLFourIconSetValues.FourTrafficLights:
                    vIcon1 = SLIconValues.BlackCircleWithBorder;
                    vIcon2 = SLIconValues.RedCircleWithBorder;
                    vIcon3 = SLIconValues.YellowCircle;
                    vIcon4 = SLIconValues.GreenCircle;
                    break;
            }

            Value2 = "25";
            Value3 = "50";
            Value4 = "75";

            Type2 = SLConditionalFormatRangeValues.Percent;
            Type3 = SLConditionalFormatRangeValues.Percent;
            Type4 = SLConditionalFormatRangeValues.Percent;
        }

        internal SLFourIconSetValues IconSetType { get; set; }

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
    }
}