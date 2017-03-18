using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.workbook
{
    /// <summary>
    ///     This simulates the DocumentFormat.OpenXml.Spreadsheet.DefinedName class.
    /// </summary>
    public class SLDefinedName
    {
        internal SLDefinedName(string Name)
        {
            Text = string.Empty;
            this.Name = Name;
            SetAllNull();
        }

        /// <summary>
        ///     The text of the defined name.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        ///     The name of the defined name. Names starting with "_xlnm" are reserved.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///     User comment.
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        ///     Custom menu text.
        /// </summary>
        public string CustomMenu { get; set; }

        /// <summary>
        ///     Description text.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        ///     Help topic for display.
        /// </summary>
        public string Help { get; set; }

        /// <summary>
        ///     Status bar text.
        /// </summary>
        public string StatusBar { get; set; }

        /// <summary>
        ///     The sheet index (0-based indexing) that's the scope of the defined name. If null, the defined name applies to the
        ///     entire spreadsheet.
        /// </summary>
        public uint? LocalSheetId { get; set; }

        /// <summary>
        ///     Specifies if the defined name is hidden in the user interface. The default value is false.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        ///     Specifies if the defined name refers to a user-defined function. The default value is false.
        /// </summary>
        public bool? Function { get; set; }

        /// <summary>
        ///     Specifies if the defined name is related to an external function, command or executable code. The default value is
        ///     false.
        /// </summary>
        public bool? VbProcedure { get; set; }

        /// <summary>
        ///     Specifies if the defined name is related to an external function, command or executable code. The default value is
        ///     false.
        /// </summary>
        public bool? Xlm { get; set; }

        /// <summary>
        ///     Specifies the function group index if the defined name refers to a function. Refer to Open XML specifications for
        ///     the meaning of the values. For example, 1 is for "Financial" and 2 is for "Date and Time".
        /// </summary>
        public uint? FunctionGroupId { get; set; }

        /// <summary>
        ///     Specifies the keyboard shortcut for the defined name.
        /// </summary>
        public string ShortcutKey { get; set; }

        /// <summary>
        ///     Specifies if the defined name is included in a spreadsheet that's published or rendered on a web or application
        ///     server. The default value is false.
        /// </summary>
        public bool? PublishToServer { get; set; }

        /// <summary>
        ///     Specifies that the defined name is used as a parameter of a spreadsheet that's published or rendered on a web or
        ///     application server. The default value is false.
        /// </summary>
        public bool? WorkbookParameter { get; set; }

        private void SetAllNull()
        {
            Comment = null;
            CustomMenu = null;
            Description = null;
            Help = null;
            StatusBar = null;
            LocalSheetId = null;
            Hidden = null;
            Function = null;
            VbProcedure = null;
            Xlm = null;
            FunctionGroupId = null;
            ShortcutKey = null;
            PublishToServer = null;
            WorkbookParameter = null;
        }

        internal void FromDefinedName(DefinedName dn)
        {
            SetAllNull();
            Text = dn.Text ?? string.Empty;
            Name = dn.Name.Value;
            if (dn.Comment != null) Comment = dn.Comment.Value;
            if (dn.CustomMenu != null) CustomMenu = dn.CustomMenu.Value;
            if (dn.Description != null) Description = dn.Description.Value;
            if (dn.Help != null) Help = dn.Help.Value;
            if (dn.StatusBar != null) StatusBar = dn.StatusBar.Value;
            if (dn.LocalSheetId != null) LocalSheetId = dn.LocalSheetId.Value;
            if (dn.Hidden != null) Hidden = dn.Hidden.Value;
            if (dn.Function != null) Function = dn.Function.Value;
            if (dn.VbProcedure != null) VbProcedure = dn.VbProcedure.Value;
            if (dn.Xlm != null) Xlm = dn.Xlm.Value;
            if (dn.FunctionGroupId != null) FunctionGroupId = dn.FunctionGroupId.Value;
            if (dn.ShortcutKey != null) ShortcutKey = dn.ShortcutKey.Value;
            if (dn.PublishToServer != null) PublishToServer = dn.PublishToServer.Value;
            if (dn.WorkbookParameter != null) WorkbookParameter = dn.WorkbookParameter.Value;
        }

        internal DefinedName ToDefinedName()
        {
            var dn = new DefinedName();
            dn.Text = Text;
            dn.Name = Name;
            if (Comment != null) dn.Comment = Comment;
            if (CustomMenu != null) dn.CustomMenu = CustomMenu;
            if (Description != null) dn.Description = Description;
            if (Help != null) dn.Help = Help;
            if (StatusBar != null) dn.StatusBar = StatusBar;
            if (LocalSheetId != null) dn.LocalSheetId = LocalSheetId.Value;
            if ((Hidden != null) && (Hidden != false)) dn.Hidden = Hidden.Value;
            if ((Function != null) && (Function != false)) dn.Function = Function.Value;
            if ((VbProcedure != null) && (VbProcedure != false)) dn.VbProcedure = VbProcedure.Value;
            if ((Xlm != null) && (Xlm != false)) dn.Xlm = Xlm.Value;
            if (FunctionGroupId != null) dn.FunctionGroupId = FunctionGroupId.Value;
            if (ShortcutKey != null) dn.ShortcutKey = ShortcutKey;
            if ((PublishToServer != null) && (PublishToServer != false)) dn.PublishToServer = PublishToServer.Value;
            if ((WorkbookParameter != null) && (WorkbookParameter != false))
                dn.WorkbookParameter = WorkbookParameter.Value;

            return dn;
        }

        internal SLDefinedName Clone()
        {
            var dn = new SLDefinedName(Name);
            dn.Text = Text;
            dn.Name = Name;
            dn.Comment = Comment;
            dn.CustomMenu = CustomMenu;
            dn.Description = Description;
            dn.Help = Help;
            dn.StatusBar = StatusBar;
            dn.LocalSheetId = LocalSheetId;
            dn.Hidden = Hidden;
            dn.Function = Function;
            dn.VbProcedure = VbProcedure;
            dn.Xlm = Xlm;
            dn.FunctionGroupId = FunctionGroupId;
            dn.ShortcutKey = ShortcutKey;
            dn.PublishToServer = PublishToServer;
            dn.WorkbookParameter = WorkbookParameter;

            return dn;
        }
    }
}