using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.office2010;
using SpreadsheetLightWrapper.Core.style;

namespace SpreadsheetLightWrapper.Core.conditionalformatting
{
    internal class SLConditionalFormattingRule
    {
        internal bool HasColorScale;
        internal bool HasDataBar;
        internal bool HasDifferentialFormat;
        internal bool HasIconSet;

        internal bool HasOperator;

        internal bool HasTimePeriod;

        internal SLConditionalFormattingRule()
        {
            SetAllNull();
        }

        internal List<Formula> Formulas { get; set; }
        internal SLColorScale ColorScale { get; set; }
        internal SLDataBar DataBar { get; set; }
        internal SLIconSet IconSet { get; set; }

        internal List<ConditionalFormattingRuleExtension> Extensions { get; set; }

        internal ConditionalFormatValues Type { get; set; }

        internal uint? FormatId { get; set; }
        internal SLDifferentialFormat DifferentialFormat { get; set; }

        internal int Priority { get; set; }
        internal bool StopIfTrue { get; set; }
        internal bool AboveAverage { get; set; }
        internal bool Percent { get; set; }
        internal bool Bottom { get; set; }
        internal ConditionalFormattingOperatorValues Operator { get; set; }

        internal string Text { get; set; }
        internal TimePeriodValues TimePeriod { get; set; }

        internal uint? Rank { get; set; }
        internal int? StdDev { get; set; }
        internal bool EqualAverage { get; set; }

        private void SetAllNull()
        {
            Formulas = new List<Formula>();
            ColorScale = new SLColorScale();
            HasColorScale = false;
            DataBar = new SLDataBar();
            HasDataBar = false;
            IconSet = new SLIconSet();
            HasIconSet = false;

            Extensions = new List<ConditionalFormattingRuleExtension>();

            Type = ConditionalFormatValues.DataBar;

            DifferentialFormat = new SLDifferentialFormat();
            HasDifferentialFormat = false;

            FormatId = null;
            Priority = 1;
            StopIfTrue = false;
            AboveAverage = true;
            Percent = false;
            Bottom = false;
            Operator = ConditionalFormattingOperatorValues.Equal;
            HasOperator = false;
            Text = null;
            TimePeriod = TimePeriodValues.Today;
            HasTimePeriod = false;
            Rank = null;
            StdDev = null;
            EqualAverage = false;
        }

        internal void FromConditionalFormattingRule(ConditionalFormattingRule cfr)
        {
            SetAllNull();

            if (cfr.Type != null) Type = cfr.Type.Value;
            if (cfr.FormatId != null) FormatId = cfr.FormatId.Value;
            Priority = cfr.Priority.Value;
            if (cfr.StopIfTrue != null) StopIfTrue = cfr.StopIfTrue.Value;
            if (cfr.AboveAverage != null) AboveAverage = cfr.AboveAverage.Value;
            if (cfr.Percent != null) Percent = cfr.Percent.Value;
            if (cfr.Bottom != null) Bottom = cfr.Bottom.Value;
            if (cfr.Operator != null)
            {
                Operator = cfr.Operator.Value;
                HasOperator = true;
            }
            if (cfr.Text != null) Text = cfr.Text.Value;
            if (cfr.TimePeriod != null)
            {
                TimePeriod = cfr.TimePeriod.Value;
                HasTimePeriod = true;
            }
            if (cfr.Rank != null) Rank = cfr.Rank.Value;
            if (cfr.StdDev != null) StdDev = cfr.StdDev.Value;
            if (cfr.EqualAverage != null) EqualAverage = cfr.EqualAverage.Value;

            using (var oxr = OpenXmlReader.Create(cfr))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Formula))
                    {
                        Formulas.Add((Formula) oxr.LoadCurrentElement().CloneNode(true));
                    }
                    else if (oxr.ElementType == typeof(ColorScale))
                    {
                        ColorScale = new SLColorScale();
                        ColorScale.FromColorScale((ColorScale) oxr.LoadCurrentElement());
                        HasColorScale = true;
                    }
                    else if (oxr.ElementType == typeof(DataBar))
                    {
                        DataBar = new SLDataBar();
                        DataBar.FromDataBar((DataBar) oxr.LoadCurrentElement());
                        HasDataBar = true;
                    }
                    else if (oxr.ElementType == typeof(IconSet))
                    {
                        IconSet = new SLIconSet();
                        IconSet.FromIconSet((IconSet) oxr.LoadCurrentElement());
                        HasIconSet = true;
                    }
                    else if (oxr.ElementType == typeof(ConditionalFormattingRuleExtension))
                    {
                        Extensions.Add((ConditionalFormattingRuleExtension) oxr.LoadCurrentElement().CloneNode(true));
                    }
            }
        }

        internal ConditionalFormattingRule ToConditionalFormattingRule()
        {
            var cfr = new ConditionalFormattingRule();
            cfr.Type = Type;
            if (FormatId != null) cfr.FormatId = FormatId.Value;
            cfr.Priority = Priority;
            if (StopIfTrue) cfr.StopIfTrue = StopIfTrue;
            if (!AboveAverage) cfr.AboveAverage = AboveAverage;
            if (Percent) cfr.Percent = Percent;
            if (Bottom) cfr.Bottom = Bottom;
            if (HasOperator) cfr.Operator = Operator;
            if ((Text != null) && (Text.Length > 0)) cfr.Text = Text;
            if (HasTimePeriod) cfr.TimePeriod = TimePeriod;
            if (Rank != null) cfr.Rank = Rank.Value;
            if (StdDev != null) cfr.StdDev = StdDev.Value;
            if (EqualAverage) cfr.EqualAverage = EqualAverage;

            foreach (var f in Formulas)
                cfr.Append((Formula) f.CloneNode(true));
            if (HasColorScale) cfr.Append(ColorScale.ToColorScale());
            if (HasDataBar) cfr.Append(DataBar.ToDataBar());
            if (HasIconSet) cfr.Append(IconSet.ToIconSet());

            if (Extensions.Count > 0)
            {
                var extlist = new ConditionalFormattingRuleExtensionList();
                foreach (var ext in Extensions)
                    extlist.Append((ConditionalFormattingRuleExtension) ext.CloneNode(true));
                cfr.Append(extlist);
            }

            return cfr;
        }

        internal SLConditionalFormattingRule2010 ToSLConditionalFormattingRule2010()
        {
            var cfr2010 = new SLConditionalFormattingRule2010();
            cfr2010.Type = Type;
            cfr2010.Priority = Priority;
            cfr2010.StopIfTrue = StopIfTrue;
            cfr2010.AboveAverage = AboveAverage;
            cfr2010.Percent = Percent;
            cfr2010.Bottom = Bottom;
            cfr2010.HasOperator = HasOperator;
            cfr2010.Operator = Operator;
            cfr2010.Text = Text;
            cfr2010.HasTimePeriod = HasTimePeriod;
            cfr2010.TimePeriod = TimePeriod;
            cfr2010.Rank = Rank;
            cfr2010.StandardDeviation = StdDev;
            cfr2010.EqualAverage = EqualAverage;

            foreach (var f in Formulas)
                cfr2010.Formulas.Add(new DocumentFormat.OpenXml.Office.Excel.Formula(f.Text));
            cfr2010.HasColorScale = HasColorScale;
            cfr2010.ColorScale = ColorScale.ToSLColorScale2010();
            cfr2010.HasDataBar = HasDataBar;
            cfr2010.DataBar = DataBar.ToDataBar2010();
            cfr2010.HasIconSet = HasIconSet;
            cfr2010.IconSet = IconSet.ToSLIconSet2010();

            cfr2010.HasDifferentialType = HasDifferentialFormat;
            cfr2010.DifferentialType = DifferentialFormat.Clone();

            return cfr2010;
        }

        internal SLConditionalFormattingRule Clone()
        {
            var cfr = new SLConditionalFormattingRule();

            cfr.Formulas = new List<Formula>();
            for (var i = 0; i < Formulas.Count; ++i)
                cfr.Formulas.Add((Formula) Formulas[i].CloneNode(true));

            cfr.HasColorScale = HasColorScale;
            cfr.ColorScale = ColorScale.Clone();
            cfr.HasDataBar = HasDataBar;
            cfr.DataBar = DataBar.Clone();
            cfr.HasIconSet = HasIconSet;
            cfr.IconSet = IconSet.Clone();

            cfr.Extensions = new List<ConditionalFormattingRuleExtension>();
            for (var i = 0; i < Extensions.Count; ++i)
                cfr.Extensions.Add((ConditionalFormattingRuleExtension) Extensions[i].CloneNode(true));

            cfr.Type = Type;
            cfr.FormatId = FormatId;
            cfr.HasDifferentialFormat = HasDifferentialFormat;
            cfr.DifferentialFormat = DifferentialFormat.Clone();

            cfr.Priority = Priority;
            cfr.StopIfTrue = StopIfTrue;
            cfr.AboveAverage = AboveAverage;
            cfr.Percent = Percent;
            cfr.Bottom = Bottom;

            cfr.HasOperator = HasOperator;
            cfr.Operator = Operator;
            cfr.Text = Text;
            cfr.HasTimePeriod = HasTimePeriod;
            cfr.TimePeriod = TimePeriod;
            cfr.Rank = Rank;
            cfr.StdDev = StdDev;
            cfr.EqualAverage = EqualAverage;

            return cfr;
        }
    }
}