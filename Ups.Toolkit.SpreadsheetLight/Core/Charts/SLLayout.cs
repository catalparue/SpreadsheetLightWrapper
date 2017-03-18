using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    internal class SLLayout
    {
        internal SLLayout()
        {
            LayoutTarget = C.LayoutTargetValues.Outer;
            LeftMode = C.LayoutModeValues.Factor;
            TopMode = C.LayoutModeValues.Factor;
            WidthMode = C.LayoutModeValues.Factor;
            HeightMode = C.LayoutModeValues.Factor;
            Left = null;
            Top = null;
            Width = null;
            Height = null;
        }

        internal C.LayoutTargetValues LayoutTarget { get; set; }
        internal C.LayoutModeValues LeftMode { get; set; }
        internal C.LayoutModeValues TopMode { get; set; }
        internal C.LayoutModeValues WidthMode { get; set; }
        internal C.LayoutModeValues HeightMode { get; set; }
        internal double? Left { get; set; }
        internal double? Top { get; set; }
        internal double? Width { get; set; }
        internal double? Height { get; set; }

        internal C.Layout ToLayout()
        {
            var layout = new C.Layout();

            if ((LayoutTarget != C.LayoutTargetValues.Outer)
                || (LeftMode != C.LayoutModeValues.Factor) || (TopMode != C.LayoutModeValues.Factor)
                || (WidthMode != C.LayoutModeValues.Factor) || (HeightMode != C.LayoutModeValues.Factor)
                || (Left != null) || (Top != null) || (Width != null) || (Height != null))
            {
                layout.ManualLayout = new C.ManualLayout();
                if (LayoutTarget != C.LayoutTargetValues.Outer)
                    layout.ManualLayout.LayoutTarget = new C.LayoutTarget {Val = LayoutTarget};
                if (LeftMode != C.LayoutModeValues.Factor)
                    layout.ManualLayout.LeftMode = new C.LeftMode {Val = LeftMode};
                if (TopMode != C.LayoutModeValues.Factor) layout.ManualLayout.TopMode = new C.TopMode {Val = TopMode};
                if (WidthMode != C.LayoutModeValues.Factor)
                    layout.ManualLayout.WidthMode = new C.WidthMode {Val = WidthMode};
                if (HeightMode != C.LayoutModeValues.Factor)
                    layout.ManualLayout.HeightMode = new C.HeightMode {Val = HeightMode};
                if (Left != null) layout.ManualLayout.Left = new C.Left {Val = Left.Value};
                if (Top != null) layout.ManualLayout.Top = new C.Top {Val = Top.Value};
                if (Width != null) layout.ManualLayout.Width = new C.Width {Val = Width.Value};
                if (Height != null) layout.ManualLayout.Height = new C.Height {Val = Height.Value};
            }

            return layout;
        }

        internal SLLayout Clone()
        {
            var l = new SLLayout();
            l.LayoutTarget = LayoutTarget;
            l.LeftMode = LeftMode;
            l.TopMode = TopMode;
            l.WidthMode = WidthMode;
            l.HeightMode = HeightMode;
            l.Left = Left;
            l.Top = Top;
            l.Width = Width;
            l.Height = Height;

            return l;
        }
    }
}