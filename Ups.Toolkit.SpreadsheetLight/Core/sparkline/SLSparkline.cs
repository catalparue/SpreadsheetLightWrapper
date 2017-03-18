using DocumentFormat.OpenXml.Office.Excel;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.sparkline
{
    internal class SLSparkline
    {
        internal int EndColumnIndex;
        internal int EndRowIndex;
        internal int LocationColumnIndex;
        internal int LocationRowIndex;
        internal int StartColumnIndex;
        internal int StartRowIndex;
        internal string WorksheetName;

        internal SLSparkline()
        {
            WorksheetName = string.Empty;
            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;
            LocationRowIndex = 1;
            LocationColumnIndex = 1;
        }

        internal X14.Sparkline ToSparkline()
        {
            var spk = new X14.Sparkline();

            if ((StartRowIndex == EndRowIndex) && (StartColumnIndex == EndColumnIndex))
            {
                spk.Formula = new Formula();
                spk.Formula.Text = SLTool.ToCellReference(WorksheetName, StartRowIndex, StartColumnIndex);
            }
            else
            {
                spk.Formula = new Formula();
                spk.Formula.Text = SLTool.ToCellRange(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex,
                    EndColumnIndex);
            }

            spk.ReferenceSequence = new ReferenceSequence();
            spk.ReferenceSequence.Text = SLTool.ToCellReference(LocationRowIndex, LocationColumnIndex);

            return spk;
        }

        internal SLSparkline Clone()
        {
            var spk = new SLSparkline();
            spk.WorksheetName = WorksheetName;
            spk.StartRowIndex = StartRowIndex;
            spk.StartColumnIndex = StartColumnIndex;
            spk.EndRowIndex = EndRowIndex;
            spk.EndColumnIndex = EndColumnIndex;
            spk.LocationRowIndex = LocationRowIndex;
            spk.LocationColumnIndex = LocationColumnIndex;

            return spk;
        }
    }
}