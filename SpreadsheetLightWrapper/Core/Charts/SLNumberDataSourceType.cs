using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     For BubbleSize, Minus, Plus, Values, YValues
    /// </summary>
    internal class SLNumberDataSourceType
    {
        private bool bUseNumberLiteral;
        private bool bUseNumberReference;

        internal SLNumberDataSourceType()
        {
            UseNumberReference = true;

            NumberReference = new SLNumberReference();
            NumberLiteral = new SLNumberLiteral();
        }

        internal bool UseNumberReference
        {
            get { return bUseNumberReference; }
            set
            {
                bUseNumberReference = value;
                if (value)
                {
                    bUseNumberReference = true;
                    bUseNumberLiteral = false;
                }
            }
        }

        internal SLNumberReference NumberReference { get; set; }

        internal bool UseNumberLiteral
        {
            get { return bUseNumberLiteral; }
            set
            {
                bUseNumberLiteral = value;
                if (value)
                {
                    bUseNumberReference = false;
                    bUseNumberLiteral = true;
                }
            }
        }

        internal SLNumberLiteral NumberLiteral { get; set; }

        internal C.BubbleSize ToBubbleSize()
        {
            var bs = new C.BubbleSize();
            if (UseNumberReference) bs.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) bs.NumberLiteral = NumberLiteral.ToNumberLiteral();

            return bs;
        }

        internal C.Minus ToMinus()
        {
            var minus = new C.Minus();
            if (UseNumberReference) minus.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) minus.NumberLiteral = NumberLiteral.ToNumberLiteral();

            return minus;
        }

        internal C.Plus ToPlus()
        {
            var plus = new C.Plus();
            if (UseNumberReference) plus.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) plus.NumberLiteral = NumberLiteral.ToNumberLiteral();

            return plus;
        }

        internal C.Values ToValues()
        {
            var v = new C.Values();
            if (UseNumberReference) v.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) v.NumberLiteral = NumberLiteral.ToNumberLiteral();

            return v;
        }

        internal C.YValues ToYValues()
        {
            var yv = new C.YValues();
            if (UseNumberReference) yv.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) yv.NumberLiteral = NumberLiteral.ToNumberLiteral();

            return yv;
        }

        internal SLNumberDataSourceType Clone()
        {
            var ndst = new SLNumberDataSourceType();
            ndst.bUseNumberReference = bUseNumberReference;
            ndst.NumberReference = NumberReference.Clone();
            ndst.bUseNumberLiteral = bUseNumberLiteral;
            ndst.NumberLiteral = NumberLiteral.Clone();

            return ndst;
        }
    }
}