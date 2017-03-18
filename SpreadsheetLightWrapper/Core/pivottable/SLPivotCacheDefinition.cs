using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLPivotCacheDefinition
    {
        internal bool HasTupleCache;

        internal SLPivotCacheDefinition()
        {
            SetAllNull();
        }

        internal SLCacheSource CacheSource { get; set; }
        internal List<SLCacheField> CacheFields { get; set; }
        internal List<SLCacheHierarchy> CacheHierarchies { get; set; }
        internal List<SLKpi> Kpis { get; set; }
        internal SLTupleCache TupleCache { get; set; }

        internal List<SLCalculatedItem> CalculatedItems { get; set; }
        internal List<SLCalculatedMember> CalculatedMembers { get; set; }
        internal List<SLDimension> Dimensions { get; set; }
        internal List<SLMeasureGroup> MeasureGroups { get; set; }
        internal List<SLMeasureDimensionMap> Maps { get; set; }

        internal string Id { get; set; }
        internal bool Invalid { get; set; }
        internal bool SaveData { get; set; }
        internal bool RefreshOnLoad { get; set; }
        internal bool OptimizeMemory { get; set; }
        internal bool EnableRefresh { get; set; }
        internal string RefreshedBy { get; set; }
        internal double? RefreshedDate { get; set; }
        internal bool BackgroundQuery { get; set; }
        internal uint? MissingItemsLimit { get; set; }
        internal byte CreatedVersion { get; set; }
        internal byte RefreshedVersion { get; set; }
        internal byte MinRefreshableVersion { get; set; }
        internal uint? RecordCount { get; set; }
        internal bool UpgradeOnRefresh { get; set; }
        internal bool IsTupleCache { get; set; }
        internal bool SupportSubquery { get; set; }
        internal bool SupportAdvancedDrill { get; set; }

        private void SetAllNull()
        {
            CacheSource = new SLCacheSource();
            CacheFields = new List<SLCacheField>();
            CacheHierarchies = new List<SLCacheHierarchy>();
            Kpis = new List<SLKpi>();
            HasTupleCache = false;
            TupleCache = new SLTupleCache();
            CalculatedItems = new List<SLCalculatedItem>();
            CalculatedMembers = new List<SLCalculatedMember>();
            Dimensions = new List<SLDimension>();
            MeasureGroups = new List<SLMeasureGroup>();
            Maps = new List<SLMeasureDimensionMap>();

            Id = "";
            Invalid = false;
            SaveData = true;
            RefreshOnLoad = false;
            OptimizeMemory = false;
            EnableRefresh = true;
            RefreshedBy = "";
            RefreshedDate = null;
            BackgroundQuery = false;
            MissingItemsLimit = null;

            // See SLPivotTable for similar explanation.
            CreatedVersion = 3;
            RefreshedVersion = 3;
            MinRefreshableVersion = 3;

            RecordCount = null;
            UpgradeOnRefresh = false;
            IsTupleCache = false;
            SupportSubquery = false;
            SupportAdvancedDrill = false;
        }

        internal void FromPivotCacheDefinition(PivotCacheDefinition pcd)
        {
            SetAllNull();

            if (pcd.Id != null) Id = pcd.Id.Value;
            if (pcd.Invalid != null) Invalid = pcd.Invalid.Value;
            if (pcd.SaveData != null) SaveData = pcd.SaveData.Value;
            if (pcd.RefreshOnLoad != null) RefreshOnLoad = pcd.RefreshOnLoad.Value;
            if (pcd.OptimizeMemory != null) OptimizeMemory = pcd.OptimizeMemory.Value;
            if (pcd.EnableRefresh != null) EnableRefresh = pcd.EnableRefresh.Value;
            if (pcd.RefreshedBy != null) RefreshedBy = pcd.RefreshedBy.Value;
            if (pcd.RefreshedDate != null) RefreshedDate = pcd.RefreshedDate.Value;
            if (pcd.BackgroundQuery != null) BackgroundQuery = pcd.BackgroundQuery.Value;
            if (pcd.MissingItemsLimit != null) MissingItemsLimit = pcd.MissingItemsLimit.Value;
            if (pcd.CreatedVersion != null) CreatedVersion = pcd.CreatedVersion.Value;
            if (pcd.RefreshedVersion != null) RefreshedVersion = pcd.RefreshedVersion.Value;
            if (pcd.MinRefreshableVersion != null) MinRefreshableVersion = pcd.MinRefreshableVersion.Value;
            if (pcd.RecordCount != null) RecordCount = pcd.RecordCount.Value;
            if (pcd.UpgradeOnRefresh != null) UpgradeOnRefresh = pcd.UpgradeOnRefresh.Value;
            if (pcd.IsTupleCache != null) IsTupleCache = pcd.IsTupleCache.Value;
            if (pcd.SupportSubquery != null) SupportSubquery = pcd.SupportSubquery.Value;
            if (pcd.SupportAdvancedDrill != null) SupportAdvancedDrill = pcd.SupportAdvancedDrill.Value;

            if (pcd.CacheSource != null) CacheSource.FromCacheSource(pcd.CacheSource);

            // doing one by one because it's bloody hindering awkward complicated.

            if (pcd.CacheFields != null)
            {
                SLCacheField cf;
                using (var oxr = OpenXmlReader.Create(pcd.CacheFields))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(CacheField))
                        {
                            cf = new SLCacheField();
                            cf.FromCacheField((CacheField) oxr.LoadCurrentElement());
                            CacheFields.Add(cf);
                        }
                }
            }

            if (pcd.CacheHierarchies != null)
            {
                SLCacheHierarchy ch;
                using (var oxr = OpenXmlReader.Create(pcd.CacheHierarchies))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(CacheHierarchy))
                        {
                            ch = new SLCacheHierarchy();
                            ch.FromCacheHierarchy((CacheHierarchy) oxr.LoadCurrentElement());
                            CacheHierarchies.Add(ch);
                        }
                }
            }

            if (pcd.Kpis != null)
            {
                SLKpi k;
                using (var oxr = OpenXmlReader.Create(pcd.Kpis))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(Kpi))
                        {
                            k = new SLKpi();
                            k.FromKpi((Kpi) oxr.LoadCurrentElement());
                            Kpis.Add(k);
                        }
                }
            }

            if (pcd.TupleCache != null)
            {
                TupleCache.FromTupleCache(pcd.TupleCache);
                HasTupleCache = true;
            }

            if (pcd.CalculatedItems != null)
            {
                SLCalculatedItem ci;
                using (var oxr = OpenXmlReader.Create(pcd.CalculatedItems))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(CalculatedItem))
                        {
                            ci = new SLCalculatedItem();
                            ci.FromCalculatedItem((CalculatedItem) oxr.LoadCurrentElement());
                            CalculatedItems.Add(ci);
                        }
                }
            }

            if (pcd.CalculatedMembers != null)
            {
                SLCalculatedMember cm;
                using (var oxr = OpenXmlReader.Create(pcd.CalculatedMembers))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(CalculatedMember))
                        {
                            cm = new SLCalculatedMember();
                            cm.FromCalculatedMember((CalculatedMember) oxr.LoadCurrentElement());
                            CalculatedMembers.Add(cm);
                        }
                }
            }

            if (pcd.Dimensions != null)
            {
                SLDimension d;
                using (var oxr = OpenXmlReader.Create(pcd.Dimensions))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(Dimension))
                        {
                            d = new SLDimension();
                            d.FromDimension((Dimension) oxr.LoadCurrentElement());
                            Dimensions.Add(d);
                        }
                }
            }

            if (pcd.MeasureGroups != null)
            {
                SLMeasureGroup mg;
                using (var oxr = OpenXmlReader.Create(pcd.MeasureGroups))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(MeasureGroup))
                        {
                            mg = new SLMeasureGroup();
                            mg.FromMeasureGroup((MeasureGroup) oxr.LoadCurrentElement());
                            MeasureGroups.Add(mg);
                        }
                }
            }

            if (pcd.Maps != null)
            {
                SLMeasureDimensionMap mdm;
                using (var oxr = OpenXmlReader.Create(pcd.Maps))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(MeasureDimensionMap))
                        {
                            mdm = new SLMeasureDimensionMap();
                            mdm.FromMeasureDimensionMap((MeasureDimensionMap) oxr.LoadCurrentElement());
                            Maps.Add(mdm);
                        }
                }
            }
        }

        internal PivotCacheDefinition ToPivotCacheDefinition()
        {
            var pcd = new PivotCacheDefinition();
            if ((Id != null) && (Id.Length > 0)) pcd.Id = Id;
            if (Invalid) pcd.Invalid = Invalid;
            if (SaveData != true) pcd.SaveData = SaveData;
            if (RefreshOnLoad) pcd.RefreshOnLoad = RefreshOnLoad;
            if (OptimizeMemory) pcd.OptimizeMemory = OptimizeMemory;
            if (EnableRefresh != true) pcd.EnableRefresh = EnableRefresh;
            if ((RefreshedBy != null) && (RefreshedBy.Length > 0)) pcd.RefreshedBy = RefreshedBy;
            if (RefreshedDate != null) pcd.RefreshedDate = RefreshedDate.Value;
            if (BackgroundQuery) pcd.BackgroundQuery = BackgroundQuery;
            if (MissingItemsLimit != null) pcd.MissingItemsLimit = MissingItemsLimit.Value;
            if (CreatedVersion != 0) pcd.CreatedVersion = CreatedVersion;
            if (RefreshedVersion != 0) pcd.RefreshedVersion = RefreshedVersion;
            if (MinRefreshableVersion != 0) pcd.MinRefreshableVersion = MinRefreshableVersion;
            if (RecordCount != null) pcd.RecordCount = RecordCount.Value;
            if (UpgradeOnRefresh) pcd.UpgradeOnRefresh = UpgradeOnRefresh;
            if (IsTupleCache) pcd.IsTupleCache = IsTupleCache;
            if (SupportSubquery) pcd.SupportSubquery = SupportSubquery;
            if (SupportAdvancedDrill) pcd.SupportAdvancedDrill = SupportAdvancedDrill;

            pcd.CacheSource = CacheSource.ToCacheSource();

            pcd.CacheFields = new CacheFields {Count = (uint) CacheFields.Count};
            foreach (var cf in CacheFields)
                pcd.CacheFields.Append(cf.ToCacheField());

            if (CacheHierarchies.Count > 0)
            {
                pcd.CacheHierarchies = new CacheHierarchies {Count = (uint) CacheHierarchies.Count};
                foreach (var ch in CacheHierarchies)
                    pcd.CacheHierarchies.Append(ch.ToCacheHierarchy());
            }

            if (Kpis.Count > 0)
            {
                pcd.Kpis = new Kpis {Count = (uint) Kpis.Count};
                foreach (var k in Kpis)
                    pcd.Kpis.Append(k.ToKpi());
            }

            if (HasTupleCache) pcd.TupleCache = TupleCache.ToTupleCache();

            if (CalculatedItems.Count > 0)
            {
                pcd.CalculatedItems = new CalculatedItems {Count = (uint) CalculatedItems.Count};
                foreach (var ci in CalculatedItems)
                    pcd.CalculatedItems.Append(ci.ToCalculatedItem());
            }

            if (CalculatedMembers.Count > 0)
            {
                pcd.CalculatedMembers = new CalculatedMembers {Count = (uint) CalculatedMembers.Count};
                foreach (var cm in CalculatedMembers)
                    pcd.CalculatedMembers.Append(cm.ToCalculatedMember());
            }

            if (Dimensions.Count > 0)
            {
                pcd.Dimensions = new Dimensions {Count = (uint) Dimensions.Count};
                foreach (var d in Dimensions)
                    pcd.Dimensions.Append(d.ToDimension());
            }

            if (MeasureGroups.Count > 0)
            {
                pcd.MeasureGroups = new MeasureGroups {Count = (uint) MeasureGroups.Count};
                foreach (var mg in MeasureGroups)
                    pcd.MeasureGroups.Append(mg.ToMeasureGroup());
            }

            if (Maps.Count > 0)
            {
                pcd.Maps = new Maps {Count = (uint) Maps.Count};
                foreach (var mdm in Maps)
                    pcd.Maps.Append(mdm.ToMeasureDimensionMap());
            }

            return pcd;
        }

        internal SLPivotCacheDefinition Clone()
        {
            var pcd = new SLPivotCacheDefinition();
            pcd.Id = Id;
            pcd.Invalid = Invalid;
            pcd.SaveData = SaveData;
            pcd.RefreshOnLoad = RefreshOnLoad;
            pcd.OptimizeMemory = OptimizeMemory;
            pcd.EnableRefresh = EnableRefresh;
            pcd.RefreshedBy = RefreshedBy;
            pcd.RefreshedDate = RefreshedDate.Value;
            pcd.BackgroundQuery = BackgroundQuery;
            pcd.MissingItemsLimit = MissingItemsLimit.Value;
            pcd.CreatedVersion = CreatedVersion;
            pcd.RefreshedVersion = RefreshedVersion;
            pcd.MinRefreshableVersion = MinRefreshableVersion;
            pcd.RecordCount = RecordCount.Value;
            pcd.UpgradeOnRefresh = UpgradeOnRefresh;
            pcd.IsTupleCache = IsTupleCache;
            pcd.SupportSubquery = SupportSubquery;
            pcd.SupportAdvancedDrill = SupportAdvancedDrill;

            pcd.CacheSource = CacheSource.Clone();

            pcd.CacheFields = new List<SLCacheField>();
            foreach (var cf in CacheFields)
                pcd.CacheFields.Add(cf.Clone());

            pcd.CacheHierarchies = new List<SLCacheHierarchy>();
            foreach (var ch in CacheHierarchies)
                pcd.CacheHierarchies.Add(ch.Clone());

            pcd.Kpis = new List<SLKpi>();
            foreach (var k in Kpis)
                pcd.Kpis.Add(k.Clone());

            pcd.HasTupleCache = HasTupleCache;
            pcd.TupleCache = TupleCache.Clone();

            pcd.CalculatedItems = new List<SLCalculatedItem>();
            foreach (var ci in CalculatedItems)
                pcd.CalculatedItems.Add(ci.Clone());

            pcd.CalculatedMembers = new List<SLCalculatedMember>();
            foreach (var cm in CalculatedMembers)
                pcd.CalculatedMembers.Add(cm.Clone());

            pcd.Dimensions = new List<SLDimension>();
            foreach (var d in Dimensions)
                pcd.Dimensions.Add(d.Clone());

            pcd.MeasureGroups = new List<SLMeasureGroup>();
            foreach (var mg in MeasureGroups)
                pcd.MeasureGroups.Add(mg.Clone());

            pcd.Maps = new List<SLMeasureDimensionMap>();
            foreach (var mdm in Maps)
                pcd.Maps.Add(mdm.Clone());

            return pcd;
        }
    }
}