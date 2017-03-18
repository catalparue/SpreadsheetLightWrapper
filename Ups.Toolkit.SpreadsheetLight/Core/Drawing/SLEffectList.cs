using System.Collections.Generic;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying effects such as glow, shadows, reflection and soft edges.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.EffectList class.
    /// </summary>
    public class SLEffectList
    {
        internal List<Color> listThemeColors;

        internal SLEffectList(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        internal bool HasEffectList
        {
            get
            {
                return Glow.HasGlow || (Shadow.IsInnerShadow != null)
                       || Reflection.HasReflection || SoftEdge.HasSoftEdge;
            }
        }

        // A.Blur is not accessible from Excel! Don't know what values to allow...

        internal SLGlow Glow { get; set; }

        internal SLShadowEffect Shadow { get; set; }

        internal SLReflection Reflection { get; set; }

        internal SLSoftEdge SoftEdge { get; set; }

        private void SetAllNull()
        {
            Glow = new SLGlow(listThemeColors);
            Shadow = new SLShadowEffect(listThemeColors);
            Reflection = new SLReflection();
            SoftEdge = new SLSoftEdge();
        }

        internal A.EffectList ToEffectList()
        {
            var el = new A.EffectList();

            if (Glow.HasGlow)
                el.Glow = Glow.ToGlow();

            if (Shadow.IsInnerShadow != null)
                if (Shadow.IsInnerShadow.Value)
                    el.InnerShadow = Shadow.ToInnerShadow();
                else
                    el.OuterShadow = Shadow.ToOuterShadow();

            if (Reflection.HasReflection)
                el.Reflection = Reflection.ToReflection();

            if (SoftEdge.HasSoftEdge)
                el.SoftEdge = SoftEdge.ToSoftEdge();

            return el;
        }

        internal SLEffectList Clone()
        {
            var el = new SLEffectList(listThemeColors);
            el.Glow = Glow.Clone();
            el.Shadow = Shadow.Clone();
            el.Reflection = Reflection.Clone();
            el.SoftEdge = SoftEdge.Clone();

            return el;
        }
    }
}