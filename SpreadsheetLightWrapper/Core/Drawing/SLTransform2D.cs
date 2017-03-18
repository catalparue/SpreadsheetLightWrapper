using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    internal class SLTransform2D
    {
        internal bool HasExtents;
        internal bool HasOffset;

        internal SLTransform2D()
        {
            HasOffset = false;
            Offset = new SLOffset();
            HasExtents = false;
            Extents = new SLExtents();

            Rotation = null;
            HorizontalFlip = null;
            VerticalFlip = null;
        }

        internal SLOffset Offset { get; set; }
        internal SLExtents Extents { get; set; }

        internal int? Rotation { get; set; }
        internal bool? HorizontalFlip { get; set; }
        internal bool? VerticalFlip { get; set; }

        internal A.Transform2D ToTransform2D()
        {
            var trans = new A.Transform2D();
            if (HasOffset) trans.Offset = Offset.ToOffset();
            if (HasExtents) trans.Extents = Extents.ToExtents();

            if (Rotation != null) trans.Rotation = Rotation.Value;
            if (HorizontalFlip != null) trans.HorizontalFlip = HorizontalFlip.Value;
            if (VerticalFlip != null) trans.VerticalFlip = VerticalFlip.Value;

            return trans;
        }

        internal SLTransform2D Clone()
        {
            var trans = new SLTransform2D();
            trans.HasOffset = HasOffset;
            trans.Offset = Offset.Clone();
            trans.HasExtents = HasExtents;
            trans.Extents = Extents.Clone();

            trans.Rotation = Rotation;
            trans.HorizontalFlip = HorizontalFlip;
            trans.VerticalFlip = VerticalFlip;

            return trans;
        }
    }
}