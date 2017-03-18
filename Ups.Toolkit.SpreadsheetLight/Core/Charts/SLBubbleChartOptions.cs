using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Chart customization options for bubble charts.
    /// </summary>
    public class SLBubbleChartOptions
    {
        internal uint iBubbleScale;

        /// <summary>
        ///     Initializes an instance of SLBubbleChartOptions.
        /// </summary>
        public SLBubbleChartOptions()
        {
            Bubble3D = true;
            iBubbleScale = 100;
            ShowNegativeBubbles = true;
            SizeRepresents = C.SizeRepresentsValues.Area;
        }

        /// <summary>
        ///     Specifies if the bubbles have a 3D effect.
        /// </summary>
        public bool Bubble3D { get; set; }

        /// <summary>
        ///     Scale factor in percentage of the default size, ranging from 0% to 300% (both inclusive). The default is 100%.
        /// </summary>
        public uint BubbleScale
        {
            get { return iBubbleScale; }
            set
            {
                iBubbleScale = value;
                if (iBubbleScale > 300) iBubbleScale = 300;
            }
        }

        /// <summary>
        ///     Specifies if negatively sized bubbles are shown.
        /// </summary>
        public bool ShowNegativeBubbles { get; set; }

        /// <summary>
        ///     Specifies how bubble sizes relate to the presentation of the bubbles.
        /// </summary>
        public C.SizeRepresentsValues SizeRepresents { get; set; }

        internal SLBubbleChartOptions Clone()
        {
            var bco = new SLBubbleChartOptions();
            bco.Bubble3D = Bubble3D;
            bco.iBubbleScale = iBubbleScale;
            bco.ShowNegativeBubbles = ShowNegativeBubbles;
            bco.SizeRepresents = SizeRepresents;

            return bco;
        }
    }
}