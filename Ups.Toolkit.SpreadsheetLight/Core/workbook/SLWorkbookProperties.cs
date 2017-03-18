using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.workbook
{
    internal class SLWorkbookProperties
    {
        private bool? bAllowRefreshQuery;

        private bool? bAutoCompressPictures;

        private bool? bBackupFile;

        private bool? bCheckCompatibility;

        private bool? bDate1904;

        private bool? bDateCompatibility;

        private bool? bFilterPrivacy;

        private bool? bHidePivotFieldList;

        private bool? bPromptedSolutions;

        private bool? bPublishItems;

        private bool? bRefreshAllConnections;

        private bool? bSaveExternalLinkValues;

        private bool? bShowBorderUnselectedTables;

        private bool? bShowInkAnnotation;

        private bool? bShowPivotChartFilter;

        private uint? iDefaultThemeVersion;

        private string sCodeName;

        private ObjectDisplayValues? vShowObjects;

        private UpdateLinksBehaviorValues? vUpdateLinks;

        internal SLWorkbookProperties()
        {
            SetAllNull();
        }

        internal bool HasWorkbookProperties
        {
            get
            {
                return (bDate1904 != null) || (bDateCompatibility != null) || (vShowObjects != null)
                       || (bShowBorderUnselectedTables != null) || (bFilterPrivacy != null) ||
                       (bPromptedSolutions != null)
                       || (bShowInkAnnotation != null) || (bBackupFile != null) || (bSaveExternalLinkValues != null)
                       || (vUpdateLinks != null) || (sCodeName != null) || (bHidePivotFieldList != null)
                       || (bShowPivotChartFilter != null) || (bAllowRefreshQuery != null) || (bPublishItems != null)
                       || (bCheckCompatibility != null) || (bAutoCompressPictures != null) ||
                       (bRefreshAllConnections != null)
                       || (iDefaultThemeVersion != null);
            }
        }

        internal bool Date1904
        {
            get { return bDate1904 ?? false; }
            set { bDate1904 = value; }
        }

        internal bool DateCompatibility
        {
            get { return bDateCompatibility ?? true; }
            set { bDateCompatibility = value; }
        }

        internal ObjectDisplayValues ShowObjects
        {
            get { return vShowObjects ?? ObjectDisplayValues.All; }
            set { vShowObjects = value; }
        }

        internal bool ShowBorderUnselectedTables
        {
            get { return bShowBorderUnselectedTables ?? true; }
            set { bShowBorderUnselectedTables = value; }
        }

        internal bool FilterPrivacy
        {
            get { return bFilterPrivacy ?? false; }
            set { bFilterPrivacy = value; }
        }

        internal bool PromptedSolutions
        {
            get { return bPromptedSolutions ?? false; }
            set { bPromptedSolutions = value; }
        }

        internal bool ShowInkAnnotation
        {
            get { return bShowInkAnnotation ?? true; }
            set { bShowInkAnnotation = value; }
        }

        internal bool BackupFile
        {
            get { return bBackupFile ?? false; }
            set { bBackupFile = value; }
        }

        internal bool SaveExternalLinkValues
        {
            get { return bSaveExternalLinkValues ?? true; }
            set { bSaveExternalLinkValues = value; }
        }

        internal UpdateLinksBehaviorValues UpdateLinks
        {
            get { return vUpdateLinks ?? UpdateLinksBehaviorValues.UserSet; }
            set { vUpdateLinks = value; }
        }

        internal string CodeName
        {
            get { return sCodeName ?? ""; }
            set { sCodeName = value; }
        }

        internal bool HidePivotFieldList
        {
            get { return bHidePivotFieldList ?? false; }
            set { bHidePivotFieldList = value; }
        }

        internal bool ShowPivotChartFilter
        {
            get { return bShowPivotChartFilter ?? false; }
            set { bShowPivotChartFilter = value; }
        }

        internal bool AllowRefreshQuery
        {
            get { return bAllowRefreshQuery ?? false; }
            set { bAllowRefreshQuery = value; }
        }

        internal bool PublishItems
        {
            get { return bPublishItems ?? false; }
            set { bPublishItems = value; }
        }

        internal bool CheckCompatibility
        {
            get { return bCheckCompatibility ?? false; }
            set { bCheckCompatibility = value; }
        }

        internal bool AutoCompressPictures
        {
            get { return bAutoCompressPictures ?? true; }
            set { bAutoCompressPictures = value; }
        }

        internal bool RefreshAllConnections
        {
            get { return bRefreshAllConnections ?? false; }
            set { bRefreshAllConnections = value; }
        }

        internal uint DefaultThemeVersion
        {
            get { return iDefaultThemeVersion ?? 0; }
            set { iDefaultThemeVersion = value; }
        }

        internal void SetAllNull()
        {
            bDate1904 = null;
            bDateCompatibility = null;
            vShowObjects = null;
            bShowBorderUnselectedTables = null;
            bFilterPrivacy = null;
            bPromptedSolutions = null;
            bShowInkAnnotation = null;
            bBackupFile = null;
            bSaveExternalLinkValues = null;
            vUpdateLinks = null;
            sCodeName = null;
            bHidePivotFieldList = null;
            bShowPivotChartFilter = null;
            bAllowRefreshQuery = null;
            bPublishItems = null;
            bCheckCompatibility = null;
            bAutoCompressPictures = null;
            bRefreshAllConnections = null;
            iDefaultThemeVersion = null;
        }

        internal void FromWorkbookProperties(WorkbookProperties wp)
        {
            SetAllNull();
            if (wp.Date1904 != null) Date1904 = wp.Date1904.Value;
            if (wp.DateCompatibility != null) DateCompatibility = wp.DateCompatibility.Value;
            if (wp.ShowObjects != null) ShowObjects = wp.ShowObjects.Value;
            if (wp.ShowBorderUnselectedTables != null) ShowBorderUnselectedTables = wp.ShowBorderUnselectedTables.Value;
            if (wp.FilterPrivacy != null) FilterPrivacy = wp.FilterPrivacy.Value;
            if (wp.PromptedSolutions != null) PromptedSolutions = wp.PromptedSolutions.Value;
            if (wp.ShowInkAnnotation != null) ShowInkAnnotation = wp.ShowInkAnnotation.Value;
            if (wp.BackupFile != null) BackupFile = wp.BackupFile.Value;
            if (wp.SaveExternalLinkValues != null) SaveExternalLinkValues = wp.SaveExternalLinkValues.Value;
            if (wp.UpdateLinks != null) UpdateLinks = wp.UpdateLinks.Value;
            if (wp.CodeName != null) CodeName = wp.CodeName.Value;
            if (wp.HidePivotFieldList != null) HidePivotFieldList = wp.HidePivotFieldList.Value;
            if (wp.ShowPivotChartFilter != null) ShowPivotChartFilter = wp.ShowPivotChartFilter.Value;
            if (wp.AllowRefreshQuery != null) AllowRefreshQuery = wp.AllowRefreshQuery.Value;
            if (wp.PublishItems != null) PublishItems = wp.PublishItems.Value;
            if (wp.CheckCompatibility != null) CheckCompatibility = wp.CheckCompatibility.Value;
            if (wp.AutoCompressPictures != null) AutoCompressPictures = wp.AutoCompressPictures.Value;
            if (wp.RefreshAllConnections != null) RefreshAllConnections = wp.RefreshAllConnections.Value;
            if (wp.DefaultThemeVersion != null) DefaultThemeVersion = wp.DefaultThemeVersion.Value;
        }

        internal WorkbookProperties ToWorkbookProperties()
        {
            var wp = new WorkbookProperties();
            if (bDate1904 != null) wp.Date1904 = bDate1904.Value;
            if (bDateCompatibility != null) wp.DateCompatibility = bDateCompatibility.Value;
            if (vShowObjects != null) wp.ShowObjects = vShowObjects.Value;
            if (bShowBorderUnselectedTables != null) wp.ShowBorderUnselectedTables = bShowBorderUnselectedTables.Value;
            if (bFilterPrivacy != null) wp.FilterPrivacy = bFilterPrivacy.Value;
            if (bPromptedSolutions != null) wp.PromptedSolutions = bPromptedSolutions.Value;
            if (bShowInkAnnotation != null) wp.ShowInkAnnotation = bShowInkAnnotation.Value;
            if (bBackupFile != null) wp.BackupFile = bBackupFile.Value;
            if (bSaveExternalLinkValues != null) wp.SaveExternalLinkValues = bSaveExternalLinkValues.Value;
            if (vUpdateLinks != null) wp.UpdateLinks = vUpdateLinks.Value;
            if (sCodeName != null) wp.CodeName = sCodeName;
            if (bHidePivotFieldList != null) wp.HidePivotFieldList = bHidePivotFieldList.Value;
            if (bShowPivotChartFilter != null) wp.ShowPivotChartFilter = bShowPivotChartFilter.Value;
            if (bAllowRefreshQuery != null) wp.AllowRefreshQuery = bAllowRefreshQuery.Value;
            if (bPublishItems != null) wp.PublishItems = bPublishItems.Value;
            if (bCheckCompatibility != null) wp.CheckCompatibility = bCheckCompatibility.Value;
            if (bAutoCompressPictures != null) wp.AutoCompressPictures = bAutoCompressPictures.Value;
            if (bRefreshAllConnections != null) wp.RefreshAllConnections = bRefreshAllConnections.Value;
            if (iDefaultThemeVersion != null) wp.DefaultThemeVersion = iDefaultThemeVersion.Value;

            return wp;
        }
    }
}