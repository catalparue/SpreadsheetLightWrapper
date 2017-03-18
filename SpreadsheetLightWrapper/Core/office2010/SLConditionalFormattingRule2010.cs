using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.style;
using Formula = DocumentFormat.OpenXml.Office.Excel.Formula;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLightWrapper.Core.office2010
{
    internal class SLConditionalFormattingRule2010
    {
        internal bool HasColorScale;
        internal bool HasDataBar;

        internal bool HasDifferentialType;
        internal bool HasIconSet;

        internal bool HasOperator;

        internal bool HasTimePeriod;

        internal SLConditionalFormattingRule2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingrule.aspx

        internal List<Formula> Formulas { get; set; }
        internal SLColorScale2010 ColorScale { get; set; }
        internal SLDataBar2010 DataBar { get; set; }
        internal SLIconSet2010 IconSet { get; set; }
        internal SLDifferentialFormat DifferentialType { get; set; }

        // extensions (MOAR extensions!?!?)

        internal ConditionalFormatValues? Type { get; set; }

        internal int? Priority { get; set; }
        internal bool StopIfTrue { get; set; }
        internal bool AboveAverage { get; set; }
        internal bool Percent { get; set; }
        internal bool Bottom { get; set; }
        internal ConditionalFormattingOperatorValues Operator { get; set; }

        internal string Text { get; set; }
        internal TimePeriodValues TimePeriod { get; set; }

        internal uint? Rank { get; set; }
        internal int? StandardDeviation { get; set; }
        internal bool EqualAverage { get; set; }
        internal bool ActivePresent { get; set; }
        internal string Id { get; set; }

        private void SetAllNull()
        {
            Formulas = new List<Formula>();
            ColorScale = new SLColorScale2010();
            HasColorScale = false;
            DataBar = new SLDataBar2010();
            HasDataBar = false;
            IconSet = new SLIconSet2010();
            HasIconSet = false;
            DifferentialType = new SLDifferentialFormat();
            HasDifferentialType = false;

            Type = null;

            Priority = null;
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
            StandardDeviation = null;
            EqualAverage = false;
            ActivePresent = false;
            Id = null;
        }

        internal void FromConditionalFormattingRule(X14.ConditionalFormattingRule cfr)
        {
            SetAllNull();

            if (cfr.Type != null) Type = cfr.Type.Value;
            if (cfr.Priority != null) Priority = cfr.Priority.Value;
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
            if (cfr.StandardDeviation != null) StandardDeviation = cfr.StandardDeviation.Value;
            if (cfr.EqualAverage != null) EqualAverage = cfr.EqualAverage.Value;
            if (cfr.ActivePresent != null) ActivePresent = cfr.ActivePresent.Value;
            if (cfr.Id != null) Id = cfr.Id.Value;

            using (var oxr = OpenXmlReader.Create(cfr))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Formula))
                    {
                        Formulas.Add((Formula) oxr.LoadCurrentElement().CloneNode(true));
                    }
                    else if (oxr.ElementType == typeof(X14.ColorScale))
                    {
                        ColorScale = new SLColorScale2010();
                        ColorScale.FromColorScale((X14.ColorScale) oxr.LoadCurrentElement());
                        HasColorScale = true;
                    }
                    else if (oxr.ElementType == typeof(X14.DataBar))
                    {
                        DataBar = new SLDataBar2010();
                        DataBar.FromDataBar((X14.DataBar) oxr.LoadCurrentElement());
                        HasDataBar = true;
                    }
                    else if (oxr.ElementType == typeof(X14.IconSet))
                    {
                        IconSet = new SLIconSet2010();
                        IconSet.FromIconSet((X14.IconSet) oxr.LoadCurrentElement());
                        HasIconSet = true;
                    }
                    else if (oxr.ElementType == typeof(X14.DifferentialType))
                    {
                        DifferentialType = new SLDifferentialFormat();
                        DifferentialType.FromDifferentialType((X14.DifferentialType) oxr.LoadCurrentElement());
                        HasDifferentialType = true;
                    }
            }
        }

        internal X14.ConditionalFormattingRule ToConditionalFormattingRule()
        {
            var cfr = new X14.ConditionalFormattingRule();
            if (Type != null) cfr.Type = Type.Value;
            if (Priority != null) cfr.Priority = Priority.Value;
            if (StopIfTrue) cfr.StopIfTrue = StopIfTrue;
            if (!AboveAverage) cfr.AboveAverage = AboveAverage;
            if (Percent) cfr.Percent = Percent;
            if (Bottom) cfr.Bottom = Bottom;
            if (HasOperator) cfr.Operator = Operator;
            if ((Text != null) && (Text.Length > 0)) cfr.Text = Text;
            if (HasTimePeriod) cfr.TimePeriod = TimePeriod;
            if (Rank != null) cfr.Rank = Rank.Value;
            if (StandardDeviation != null) cfr.StandardDeviation = StandardDeviation.Value;
            if (EqualAverage) cfr.EqualAverage = EqualAverage;
            if (ActivePresent) cfr.ActivePresent = ActivePresent;
            if (Id != null) cfr.Id = Id;

            foreach (var f in Formulas)
                cfr.Append((Formula) f.CloneNode(true));
            if (HasColorScale) cfr.Append(ColorScale.ToColorScale());
            if (HasDataBar) cfr.Append(DataBar.ToDataBar(Priority != null));
            if (HasIconSet) cfr.Append(IconSet.ToIconSet());
            if (HasDifferentialType) cfr.Append(DifferentialType.ToDifferentialType());

            return cfr;
        }

        internal SLConditionalFormattingRule2010 Clone()
        {
            var cfr = new SLConditionalFormattingRule2010();

            cfr.Formulas = new List<Formula>();
            for (var i = 0; i < Formulas.Count; ++i)
                cfr.Formulas.Add((Formula) Formulas[i].CloneNode(true));

            cfr.HasColorScale = HasColorScale;
            cfr.ColorScale = ColorScale.Clone();
            cfr.HasDataBar = HasDataBar;
            cfr.DataBar = DataBar.Clone();
            cfr.HasIconSet = HasIconSet;
            cfr.IconSet = IconSet.Clone();
            cfr.HasDifferentialType = HasDifferentialType;
            cfr.DifferentialType = DifferentialType.Clone();

            cfr.Type = Type;
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
            cfr.StandardDeviation = StandardDeviation;
            cfr.EqualAverage = EqualAverage;
            cfr.ActivePresent = ActivePresent;
            cfr.Id = Id;

            return cfr;
        }
    }
}