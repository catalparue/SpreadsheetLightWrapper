namespace SpreadsheetLightWrapper.Core.conditionalformatting
{
    /// <summary>
    ///     Conditional formatting options for three icon sets.
    /// </summary>
    public class SLThreeIconSetOptions
    {
        internal bool IsCustomIcon;

        internal SLIconValues vIcon1;

        internal SLIconValues vIcon2;

        internal SLIconValues vIcon3;

        /// <summary>
        ///     Initializes an instance of SLThreeIconSetOptions.
        /// </summary>
        /// <param name="IconSetType">The type of icon set.</param>
        public SLThreeIconSetOptions(SLThreeIconSetValues IconSetType)
        {
            this.IconSetType = IconSetType;
            ReverseIconOrder = false;
            ShowIconOnly = false;

            IsCustomIcon = false;

            switch (IconSetType)
            {
                case SLThreeIconSetValues.ThreeArrows:
                    vIcon1 = SLIconValues.RedDownArrow;
                    vIcon2 = SLIconValues.YellowSideArrow;
                    vIcon3 = SLIconValues.GreenUpArrow;
                    break;
                case SLThreeIconSetValues.ThreeArrowsGray:
                    vIcon1 = SLIconValues.GrayDownArrow;
                    vIcon2 = SLIconValues.GraySideArrow;
                    vIcon3 = SLIconValues.GrayUpArrow;
                    break;
                case SLThreeIconSetValues.ThreeFlags:
                    vIcon1 = SLIconValues.RedFlag;
                    vIcon2 = SLIconValues.YellowFlag;
                    vIcon3 = SLIconValues.GreenFlag;
                    break;
                case SLThreeIconSetValues.ThreeSigns:
                    vIcon1 = SLIconValues.RedDiamond;
                    vIcon2 = SLIconValues.YellowTriangle;
                    vIcon3 = SLIconValues.GreenCircle;
                    break;
                case SLThreeIconSetValues.ThreeStars:
                    vIcon1 = SLIconValues.SilverStar;
                    vIcon2 = SLIconValues.HalfGoldStar;
                    vIcon3 = SLIconValues.GoldStar;
                    break;
                case SLThreeIconSetValues.ThreeSymbols:
                    vIcon1 = SLIconValues.RedCrossSymbol;
                    vIcon2 = SLIconValues.YellowExclamationSymbol;
                    vIcon3 = SLIconValues.GreenCheckSymbol;
                    break;
                case SLThreeIconSetValues.ThreeSymbols2:
                    vIcon1 = SLIconValues.RedCross;
                    vIcon2 = SLIconValues.YellowExclamation;
                    vIcon3 = SLIconValues.GreenCheck;
                    break;
                case SLThreeIconSetValues.ThreeTrafficLights1:
                    vIcon1 = SLIconValues.RedCircleWithBorder;
                    vIcon2 = SLIconValues.YellowCircle;
                    vIcon3 = SLIconValues.GreenCircle;
                    break;
                case SLThreeIconSetValues.ThreeTrafficLights2:
                    vIcon1 = SLIconValues.RedTrafficLight;
                    vIcon2 = SLIconValues.YellowTrafficLight;
                    vIcon3 = SLIconValues.GreenTrafficLight;
                    break;
                case SLThreeIconSetValues.ThreeTriangles:
                    vIcon1 = SLIconValues.RedDownTriangle;
                    vIcon2 = SLIconValues.YellowDash;
                    vIcon3 = SLIconValues.GreenUpTriangle;
                    break;
            }

            GreaterThanOrEqual2 = true;
            GreaterThanOrEqual3 = true;

            Value2 = "33";
            Value3 = "67";

            Type2 = SLConditionalFormatRangeValues.Percent;
            Type3 = SLConditionalFormatRangeValues.Percent;
        }

        internal SLThreeIconSetValues IconSetType { get; set; }

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
        ///     The 2nd range value.
        /// </summary>
        public string Value2 { get; set; }

        /// <summary>
        ///     The 3rd range value.
        /// </summary>
        public string Value3 { get; set; }

        /// <summary>
        ///     The conditional format type for the 2nd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type2 { get; set; }

        /// <summary>
        ///     The conditional format type for the 3rd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type3 { get; set; }
    }
}