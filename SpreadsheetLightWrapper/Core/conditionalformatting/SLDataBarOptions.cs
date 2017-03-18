using System.Collections.Generic;
using SpreadsheetLightWrapper.Core.style;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLightWrapper.Core.conditionalformatting
{
    /// <summary>
    ///     Conditional formatting options for data bars.
    /// </summary>
    public class SLDataBarOptions
    {
        internal bool bBorder;

        internal bool bGradient;

        internal bool bNegativeBarBorderColorSameAsPositive;

        internal bool bNegativeBarColorSameAsPositive;
        internal bool Is2010;

        internal X14.DataBarAxisPositionValues vAxisPosition;

        internal X14.DataBarDirectionValues vDirection;

        internal SLConditionalFormatAutoMinMaxValues vMaximumType;

        internal SLConditionalFormatAutoMinMaxValues vMinimumType;

        /// <summary>
        ///     Initializes an instance of SLDataBarOptions.
        /// </summary>
        public SLDataBarOptions()
        {
            InitialiseDataBarOptions(SLConditionalFormatDataBarValues.Blue, true);
        }

        /// <summary>
        ///     Initializes an instance of SLDataBarOptions.
        /// </summary>
        /// <param name="DataBar">Built-in data bar type.</param>
        public SLDataBarOptions(SLConditionalFormatDataBarValues DataBar)
        {
            InitialiseDataBarOptions(DataBar, true);
        }

        /// <summary>
        ///     Initializes an instance of SLDataBarOptions.
        /// </summary>
        /// <param name="Is2010Default">True if Excel 2010 specific data bar is to be used. False otherwise.</param>
        public SLDataBarOptions(bool Is2010Default)
        {
            InitialiseDataBarOptions(SLConditionalFormatDataBarValues.Blue, Is2010Default);
        }

        /// <summary>
        ///     Initializes an instance of SLDataBarOptions.
        /// </summary>
        /// <param name="DataBar">Built-in data bar type.</param>
        /// <param name="Is2010Default">True if Excel 2010 specific data bar is to be used. False otherwise.</param>
        public SLDataBarOptions(SLConditionalFormatDataBarValues DataBar, bool Is2010Default)
        {
            InitialiseDataBarOptions(DataBar, Is2010Default);
        }

        /// <summary>
        ///     The conditional format type for the minimum value. If "Automatic" is used, Excel 2010 specific data bars will be
        ///     used.
        /// </summary>
        public SLConditionalFormatAutoMinMaxValues MinimumType
        {
            get { return vMinimumType; }
            set
            {
                vMinimumType = value;
                if (vMinimumType == SLConditionalFormatAutoMinMaxValues.Automatic) Is2010 = true;
            }
        }

        /// <summary>
        ///     The minimum value.
        /// </summary>
        public string MinimumValue { get; set; }

        /// <summary>
        ///     The conditional format type for the maximum value. If "Automatic" is used, Excel 2010 specific data bars will be
        ///     used.
        /// </summary>
        public SLConditionalFormatAutoMinMaxValues MaximumType
        {
            get { return vMaximumType; }
            set
            {
                vMaximumType = value;
                if (vMaximumType == SLConditionalFormatAutoMinMaxValues.Automatic) Is2010 = true;
            }
        }

        /// <summary>
        ///     The maximum value.
        /// </summary>
        public string MaximumValue { get; set; }

        /// <summary>
        ///     The fill color.
        /// </summary>
        public SLColor FillColor { get; set; }

        /// <summary>
        ///     The border color.
        /// </summary>
        public SLColor BorderColor { get; set; }

        /// <summary>
        ///     The fill color for negative values.
        /// </summary>
        public SLColor NegativeFillColor { get; set; }

        /// <summary>
        ///     The border color for negative values.
        /// </summary>
        public SLColor NegativeBorderColor { get; set; }

        /// <summary>
        ///     The axis color.
        /// </summary>
        public SLColor AxisColor { get; set; }

        /// <summary>
        ///     The minimum length of the data bar as a percentage of the cell width. The default value is 10.
        /// </summary>
        public uint MinLength { get; set; }

        /// <summary>
        ///     The maximum length of the data bar as a percentage of the cell width. The default value is 90. It is recommended to
        ///     keep this to a maximum (haha) of 100.
        /// </summary>
        public uint MaxLength { get; set; }

        /// <summary>
        ///     Specifies if only the data bar is shown. Set to false to show both data bar and value.
        /// </summary>
        public bool ShowBarOnly { get; set; }

        /// <summary>
        ///     Specifies if there's a border. This is an Excel 2010 specific feature.
        /// </summary>
        public bool Border
        {
            get { return bBorder; }
            set
            {
                bBorder = value;
                Is2010 = true;
            }
        }

        /// <summary>
        ///     Specifies if the fill color has a gradient. This is an Excel 2010 specific feature.
        /// </summary>
        public bool Gradient
        {
            get { return bGradient; }
            set
            {
                bGradient = value;
                Is2010 = true;
            }
        }

        /// <summary>
        ///     The bar direction. This is an Excel 2010 specific feature.
        /// </summary>
        public X14.DataBarDirectionValues Direction
        {
            get { return vDirection; }
            set
            {
                vDirection = value;
                Is2010 = true;
            }
        }

        /// <summary>
        ///     Specifies if the fill color for negative values is the same as the positive one. This is an Excel 2010 specific
        ///     feature.
        /// </summary>
        public bool NegativeBarColorSameAsPositive
        {
            get { return bNegativeBarColorSameAsPositive; }
            set
            {
                bNegativeBarColorSameAsPositive = value;
                Is2010 = true;
            }
        }

        /// <summary>
        ///     Specifies if the border color for negative values is the same as the positive one. This is an Excel 2010 specific
        ///     feature.
        /// </summary>
        public bool NegativeBarBorderColorSameAsPositive
        {
            get { return bNegativeBarBorderColorSameAsPositive; }
            set
            {
                bNegativeBarBorderColorSameAsPositive = value;
                Is2010 = true;
            }
        }

        /// <summary>
        ///     Specifies the axis position. This is an Excel 2010 specific feature.
        /// </summary>
        public X14.DataBarAxisPositionValues AxisPosition
        {
            get { return vAxisPosition; }
            set
            {
                vAxisPosition = value;
                Is2010 = true;
            }
        }

        private void InitialiseDataBarOptions(SLConditionalFormatDataBarValues DataBar, bool Is2010Default)
        {
            Is2010 = Is2010Default;

            FillColor = new SLColor(new List<Color>(), new List<Color>());
            BorderColor = new SLColor(new List<Color>(), new List<Color>());
            NegativeFillColor = new SLColor(new List<Color>(), new List<Color>());
            NegativeBorderColor = new SLColor(new List<Color>(), new List<Color>());
            AxisColor = new SLColor(new List<Color>(), new List<Color>());

            switch (DataBar)
            {
                case SLConditionalFormatDataBarValues.Blue:
                    FillColor.Color = Color.FromArgb(0xFF, 0x63, 0x8E, 0xC6);
                    BorderColor.Color = Color.FromArgb(0xFF, 0x63, 0x8E, 0xC6);
                    break;
                case SLConditionalFormatDataBarValues.Green:
                    FillColor.Color = Color.FromArgb(0xFF, 0x63, 0xC3, 0x84);
                    BorderColor.Color = Color.FromArgb(0xFF, 0x63, 0xC3, 0x84);
                    break;
                case SLConditionalFormatDataBarValues.Red:
                    FillColor.Color = Color.FromArgb(0xFF, 0xFF, 0x55, 0x5A);
                    BorderColor.Color = Color.FromArgb(0xFF, 0xFF, 0x55, 0x5A);
                    break;
                case SLConditionalFormatDataBarValues.Orange:
                    FillColor.Color = Color.FromArgb(0xFF, 0xFF, 0xB6, 0x28);
                    BorderColor.Color = Color.FromArgb(0xFF, 0xFF, 0xB6, 0x28);
                    break;
                case SLConditionalFormatDataBarValues.LightBlue:
                    FillColor.Color = Color.FromArgb(0xFF, 0x00, 0x8A, 0xEF);
                    BorderColor.Color = Color.FromArgb(0xFF, 0x00, 0x8A, 0xEF);
                    break;
                case SLConditionalFormatDataBarValues.Purple:
                    FillColor.Color = Color.FromArgb(0xFF, 0xD6, 0x00, 0x7B);
                    BorderColor.Color = Color.FromArgb(0xFF, 0xD6, 0x00, 0x7B);
                    break;
            }

            NegativeFillColor.Color = Color.FromArgb(0xFF, 0xFF, 0x00, 0x00);
            NegativeBorderColor.Color = Color.FromArgb(0xFF, 0xFF, 0x00, 0x00);
            AxisColor.Color = Color.FromArgb(0xFF, 0x00, 0x00, 0x00);

            if (Is2010Default)
            {
                vMinimumType = SLConditionalFormatAutoMinMaxValues.Automatic;
                MinimumValue = string.Empty;
                vMaximumType = SLConditionalFormatAutoMinMaxValues.Automatic;
                MaximumValue = string.Empty;
                MinLength = 0;
                MaxLength = 100;
            }
            else
            {
                vMinimumType = SLConditionalFormatAutoMinMaxValues.Value;
                MinimumValue = string.Empty;
                vMaximumType = SLConditionalFormatAutoMinMaxValues.Value;
                MaximumValue = string.Empty;
                MinLength = 10;
                MaxLength = 90;
            }

            ShowBarOnly = false;
            bBorder = false;
            bGradient = false;
            vDirection = X14.DataBarDirectionValues.Context;
            bNegativeBarColorSameAsPositive = false;
            bNegativeBarBorderColorSameAsPositive = true;
            vAxisPosition = X14.DataBarAxisPositionValues.Automatic;
        }
    }
}