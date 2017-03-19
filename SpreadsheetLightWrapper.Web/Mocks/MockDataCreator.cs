using System;
using System.Collections.Generic;
using System.Data;

namespace SpreadsheetLightWrapper.Web.Mocks
{
    /// ===========================================================================================
    /// <summary>
    ///     Mock Data Generator Related, Unrelated & Partially Related Tables
    /// </summary>
    /// ===========================================================================================
    public class MockDataCreator
    {
        private readonly Associates _associates;
        // Members
        private readonly DataSet _dataSet;
        private readonly Directors _directors;
        private readonly Managers _managers;
        private readonly TeamLeads _teamLeads;

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Base Constructor
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public MockDataCreator()
        {
            try
            {
                _dataSet = new DataSet();
                _directors = new Directors();
                _managers = new Managers();
                _teamLeads = new TeamLeads();
                _associates = new Associates();
            }
            catch (Exception ex)
            {

            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     For multiple-tables that are related in a Parent-Child configuration there must always
        ///     be a primary key -> foreign key relation, otherwise the related child table will be skipped
        ///     ** Note: As an experiment comment one or more of the relations and test the result
        /// </summary>
        /// <returns>DataSet</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataSet CreateRelatedGroupedDataSet()
        {
            try
            {
                _dataSet.Tables.Add(_directors.CreateDirectors());
                _dataSet.Tables.Add(_managers.CreateManagers());
                _dataSet.Tables.Add(_teamLeads.CreateTeamLeads());
                _dataSet.Tables.Add(_associates.CreateAssociates());

                _dataSet.Relations.Add("FK_Managers_Directors",
                    _dataSet.Tables["Directors"].Columns["DID"],
                    _dataSet.Tables["Managers"].Columns["DID"]);

                _dataSet.Relations.Add("FK_TeamLeads_Managers",
                    _dataSet.Tables["Managers"].Columns["MID"],
                    _dataSet.Tables["TeamLeads"].Columns["MID"]);

                _dataSet.Relations.Add("FK_Associates_TeamLeads",
                    _dataSet.Tables["TeamLeads"].Columns["TLID"],
                    _dataSet.Tables["Associates"].Columns["TLID"]);

                return _dataSet;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     As with the related table function there has to be relation between the two related tables
        ///     in this group
        /// </summary>
        /// <returns>DataSet</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataSet CreatePartiallyRelatedGroupedDataSet()
        {
            try
            {
                _dataSet.Tables.Add(_directors.CreateDirectors());
                _dataSet.Tables.Add(_managers.CreateManagers());
                _dataSet.Tables.Add(_teamLeads.CreateTeamLeads());
                _dataSet.Tables.Add(_associates.CreateAssociates());

                _dataSet.Relations.Add("FK_Associates_TeamLeads",
                    _dataSet.Tables["TeamLeads"].Columns["TLID"],
                    _dataSet.Tables["Associates"].Columns["TLID"]);

                return _dataSet;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     For multiple-tables that are unrelated
        /// </summary>
        /// <returns>DataSet</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataSet CreateUnrelatedUngroupedDataSet()
        {
            try
            {
                _dataSet.Tables.Add(_directors.CreateDirectors());
                _dataSet.Tables.Add(_managers.CreateManagers());
                _dataSet.Tables.Add(_teamLeads.CreateTeamLeads());
                _dataSet.Tables.Add(_associates.CreateAssociates());

                return _dataSet;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This example has four tables with two that are related, Managers & TeamLeads.
        ///     The unrelated tables get their own sheets, the two related go on one sheet.
        /// </summary>
        /// <returns>DataSet</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataSet CreatePartiallyRelatedGroupedDataSetVer2()
        {
            try
            {
                _dataSet.Tables.Add(_directors.CreateDirectors());
                _dataSet.Tables.Add(_managers.CreateManagers());
                _dataSet.Tables.Add(_teamLeads.CreateTeamLeads());
                _dataSet.Tables.Add(_associates.CreateAssociates());

                _dataSet.Relations.Add("FK_TeamLeads_Managers",
                     _dataSet.Tables["Managers"].Columns["MID"],
                     _dataSet.Tables["TeamLeads"].Columns["MID"]);

                return _dataSet;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This example has four tables with three that are related, Directors, Managers & TeamLeads.  
        ///     The unrelated table Associates will get its own sheet, the three related will be grouped 
        ///     on one sheet.
        /// </summary>
        /// <returns>DataSet</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataSet CreatePartiallyRelatedGroupedDataSetVer3()
        {
            try
            {
                _dataSet.Tables.Add(_directors.CreateDirectors());
                _dataSet.Tables.Add(_managers.CreateManagers());
                _dataSet.Tables.Add(_teamLeads.CreateTeamLeads());
                _dataSet.Tables.Add(_associates.CreateAssociates());

                _dataSet.Relations.Add("FK_Managers_Directors",
                    _dataSet.Tables["Directors"].Columns["DID"],
                    _dataSet.Tables["Managers"].Columns["DID"]);

                _dataSet.Relations.Add("FK_TeamLeads_Managers",
                     _dataSet.Tables["Managers"].Columns["MID"],
                     _dataSet.Tables["TeamLeads"].Columns["MID"]);

                return _dataSet;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Destroy all the left over objects
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        ~MockDataCreator()
        {
            try
            {
                _dataSet.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
    }
}