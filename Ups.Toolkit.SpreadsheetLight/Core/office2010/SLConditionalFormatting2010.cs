using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.Excel;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.worksheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.office2010
{
    internal class SLConditionalFormatting2010
    {
        internal SLConditionalFormatting2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformatting.aspx

        internal List<SLConditionalFormattingRule2010> Rules { get; set; }
        internal List<SLCellPointRange> ReferenceSequence { get; set; }

        // extensions?

        internal bool Pivot { get; set; }

        private void SetAllNull()
        {
            Rules = new List<SLConditionalFormattingRule2010>();
            ReferenceSequence = new List<SLCellPointRange>();
            Pivot = false;
        }

        internal void FromConditionalFormatting(X14.ConditionalFormatting cf)
        {
            SetAllNull();

            if (cf.Pivot != null) Pivot = cf.Pivot.Value;

            using (var oxr = OpenXmlReader.Create(cf))
            {
                while (oxr.Read())
                {
                    SLConditionalFormattingRule2010 cfr;
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingRule))
                    {
                        cfr = new SLConditionalFormattingRule2010();
                        cfr.FromConditionalFormattingRule((X14.ConditionalFormattingRule) oxr.LoadCurrentElement());
                        Rules.Add(cfr);
                    }
                    else if (oxr.ElementType == typeof(ReferenceSequence))
                    {
                        var refseq = (ReferenceSequence) oxr.LoadCurrentElement();
                        ReferenceSequence = SLTool.TranslateRefSeqToCellPointRange(refseq);
                    }
                }
            }
        }

        internal X14.ConditionalFormatting ToConditionalFormatting()
        {
            var cf = new X14.ConditionalFormatting();
            // otherwise xm:f and xm:seqref becomes xne:f and xne:seqref
            cf.AddNamespaceDeclaration("xm", SLConstants.NamespaceXm);
            // how come sparklines don't need explicit namespace declarations?

            if (Pivot) cf.Pivot = Pivot;

            int i;
            for (i = 0; i < Rules.Count; ++i)
                cf.Append(Rules[i].ToConditionalFormattingRule());

            if (ReferenceSequence.Count > 0)
                cf.Append(new ReferenceSequence(SLTool.TranslateCellPointRangeToRefSeq(ReferenceSequence)));

            return cf;
        }

        internal SLConditionalFormatting2010 Clone()
        {
            var cf = new SLConditionalFormatting2010();

            int i;
            cf.Rules = new List<SLConditionalFormattingRule2010>();
            for (i = 0; i < Rules.Count; ++i)
                cf.Rules.Add(Rules[i].Clone());

            cf.ReferenceSequence = new List<SLCellPointRange>();
            SLCellPointRange cpr;
            for (i = 0; i < ReferenceSequence.Count; ++i)
            {
                cpr = new SLCellPointRange(ReferenceSequence[i].StartRowIndex, ReferenceSequence[i].StartColumnIndex,
                    ReferenceSequence[i].EndRowIndex, ReferenceSequence[i].EndColumnIndex);
                cf.ReferenceSequence.Add(cpr);
            }

            cf.Pivot = Pivot;

            return cf;
        }
    }
}