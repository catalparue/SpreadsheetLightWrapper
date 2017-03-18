using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     For CategoryAxisData and XValues
    /// </summary>
    internal class SLAxisDataSourceType
    {
        protected bool bUseMultiLevelStringReference;

        protected bool bUseNumberLiteral;

        protected bool bUseNumberReference;

        protected bool bUseStringLiteral;

        protected bool bUseStringReference;

        internal SLAxisDataSourceType()
        {
            UseStringReference = true;

            MultiLevelStringReference = new SLMultiLevelStringReference();
            NumberReference = new SLNumberReference();
            NumberLiteral = new SLNumberLiteral();
            StringReference = new SLStringReference();
            StringLiteral = new SLStringLiteral();
        }

        internal bool UseMultiLevelStringReference
        {
            get { return bUseMultiLevelStringReference; }
            set
            {
                bUseMultiLevelStringReference = value;
                if (value)
                {
                    bUseMultiLevelStringReference = true;
                    bUseNumberReference = false;
                    bUseNumberLiteral = false;
                    bUseStringReference = false;
                    bUseStringLiteral = false;
                }
            }
        }

        internal SLMultiLevelStringReference MultiLevelStringReference { get; set; }

        internal bool UseNumberReference
        {
            get { return bUseNumberReference; }
            set
            {
                bUseNumberReference = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = true;
                    bUseNumberLiteral = false;
                    bUseStringReference = false;
                    bUseStringLiteral = false;
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
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = false;
                    bUseNumberLiteral = true;
                    bUseStringReference = false;
                    bUseStringLiteral = false;
                }
            }
        }

        internal SLNumberLiteral NumberLiteral { get; set; }

        internal bool UseStringReference
        {
            get { return bUseStringReference; }
            set
            {
                bUseStringReference = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = false;
                    bUseNumberLiteral = false;
                    bUseStringReference = true;
                    bUseStringLiteral = false;
                }
            }
        }

        internal SLStringReference StringReference { get; set; }

        internal bool UseStringLiteral
        {
            get { return bUseStringLiteral; }
            set
            {
                bUseStringLiteral = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = false;
                    bUseNumberLiteral = false;
                    bUseStringReference = false;
                    bUseStringLiteral = true;
                }
            }
        }

        internal SLStringLiteral StringLiteral { get; set; }

        internal C.CategoryAxisData ToCategoryAxisData()
        {
            var cad = new C.CategoryAxisData();
            if (UseMultiLevelStringReference)
                cad.MultiLevelStringReference = MultiLevelStringReference.ToMultiLevelStringReference();
            if (UseNumberReference) cad.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) cad.NumberLiteral = NumberLiteral.ToNumberLiteral();
            if (UseStringReference) cad.StringReference = StringReference.ToStringReference();
            if (UseStringLiteral) cad.StringLiteral = StringLiteral.ToStringLiteral();

            return cad;
        }

        internal C.XValues ToXValues()
        {
            var xv = new C.XValues();
            if (UseMultiLevelStringReference)
                xv.MultiLevelStringReference = MultiLevelStringReference.ToMultiLevelStringReference();
            if (UseNumberReference) xv.NumberReference = NumberReference.ToNumberReference();
            if (UseNumberLiteral) xv.NumberLiteral = NumberLiteral.ToNumberLiteral();
            if (UseStringReference) xv.StringReference = StringReference.ToStringReference();
            if (UseStringLiteral) xv.StringLiteral = StringLiteral.ToStringLiteral();

            return xv;
        }

        internal SLAxisDataSourceType Clone()
        {
            var adst = new SLAxisDataSourceType();
            adst.bUseMultiLevelStringReference = bUseMultiLevelStringReference;
            adst.bUseNumberLiteral = bUseNumberLiteral;
            adst.bUseNumberReference = bUseNumberReference;
            adst.bUseStringLiteral = bUseStringLiteral;
            adst.bUseStringReference = bUseStringReference;

            adst.MultiLevelStringReference = MultiLevelStringReference.Clone();
            adst.NumberLiteral = NumberLiteral.Clone();
            adst.NumberReference = NumberReference.Clone();
            adst.StringLiteral = StringLiteral.Clone();
            adst.StringReference = StringReference.Clone();

            return adst;
        }
    }
}