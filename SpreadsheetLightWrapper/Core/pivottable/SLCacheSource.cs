using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLCacheSource
    {
        /// <summary>
        ///     If true, use worksheet. If false, use consolidation. If null, use extension list.
        /// </summary>
        internal bool? IsWorksheetSource;

        internal SLCacheSource()
        {
            SetAllNull();
        }

        // for WorksheetSource
        internal string WorksheetSourceReference { get; set; }
        internal string WorksheetSourceName { get; set; }
        internal string WorksheetSourceSheet { get; set; }
        internal string WorksheetSourceId { get; set; }

        internal SLConsolidation Consolidation { get; set; }
        internal CacheSourceExtensionList ExtensionList { get; set; }

        internal SourceValues Type { get; set; }
        internal uint ConnectionId { get; set; }

        private void SetAllNull()
        {
            IsWorksheetSource = true;

            WorksheetSourceReference = "";
            WorksheetSourceName = "";
            WorksheetSourceSheet = "";
            WorksheetSourceId = "";

            Consolidation = new SLConsolidation();
            ExtensionList = null;

            Type = SourceValues.Worksheet;
            ConnectionId = 0;
        }

        internal void FromCacheSource(CacheSource cs)
        {
            SetAllNull();

            if (cs.Type != null) Type = cs.Type.Value;
            if (cs.ConnectionId != null) ConnectionId = cs.ConnectionId.Value;

            if (cs.WorksheetSource != null)
            {
                if (cs.WorksheetSource.Reference != null) WorksheetSourceReference = cs.WorksheetSource.Reference.Value;
                if (cs.WorksheetSource.Name != null) WorksheetSourceName = cs.WorksheetSource.Name.Value;
                if (cs.WorksheetSource.Sheet != null) WorksheetSourceSheet = cs.WorksheetSource.Sheet.Value;
                if (cs.WorksheetSource.Id != null) WorksheetSourceId = cs.WorksheetSource.Id.Value;
                IsWorksheetSource = true;
            }
            else if (cs.Consolidation != null)
            {
                Consolidation.FromConsolidation(cs.Consolidation);
                IsWorksheetSource = false;
            }
            else if (cs.CacheSourceExtensionList != null)
            {
                ExtensionList = (CacheSourceExtensionList) cs.CacheSourceExtensionList.CloneNode(true);
                IsWorksheetSource = null;
            }
        }

        internal CacheSource ToCacheSource()
        {
            var cs = new CacheSource();

            cs.Type = Type;
            if (ConnectionId != 0) cs.ConnectionId = ConnectionId;

            if (IsWorksheetSource != null)
            {
                if (IsWorksheetSource.Value)
                {
                    cs.WorksheetSource = new WorksheetSource();
                    if ((WorksheetSourceReference != null) && (WorksheetSourceReference.Length > 0))
                        cs.WorksheetSource.Reference = WorksheetSourceReference;
                    if ((WorksheetSourceName != null) && (WorksheetSourceName.Length > 0))
                        cs.WorksheetSource.Name = WorksheetSourceName;
                    if ((WorksheetSourceSheet != null) && (WorksheetSourceSheet.Length > 0))
                        cs.WorksheetSource.Sheet = WorksheetSourceSheet;
                    if ((WorksheetSourceId != null) && (WorksheetSourceId.Length > 0))
                        cs.WorksheetSource.Id = WorksheetSourceId;
                }
                else
                {
                    cs.Consolidation = Consolidation.ToConsolidation();
                }
            }
            else
            {
                if (ExtensionList != null)
                    cs.CacheSourceExtensionList = (CacheSourceExtensionList) ExtensionList.CloneNode(true);
            }

            return cs;
        }

        internal SLCacheSource Clone()
        {
            var cs = new SLCacheSource();
            cs.Type = Type;
            cs.ConnectionId = ConnectionId;

            cs.IsWorksheetSource = IsWorksheetSource;

            cs.WorksheetSourceReference = WorksheetSourceReference;
            cs.WorksheetSourceName = WorksheetSourceName;
            cs.WorksheetSourceSheet = WorksheetSourceSheet;
            cs.WorksheetSourceId = WorksheetSourceId;

            cs.Consolidation = Consolidation.Clone();

            if (ExtensionList != null) cs.ExtensionList = (CacheSourceExtensionList) ExtensionList.CloneNode(true);

            return cs;
        }
    }
}