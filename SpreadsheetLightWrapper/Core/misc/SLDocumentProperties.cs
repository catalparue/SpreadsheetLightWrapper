namespace SpreadsheetLightWrapper.Core.misc
{
    /// <summary>
    ///     Encapsulates properties and methods for setting spreadsheet document properties.
    /// </summary>
    public class SLDocumentProperties
    {
        internal SLDocumentProperties()
        {
            SetAllNull();
        }

        /// <summary>
        ///     The category of the document.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        ///     The status of the content.
        /// </summary>
        public string ContentStatus { get; set; }

        internal string Created { get; set; }

        /// <summary>
        ///     The creator of the document.
        /// </summary>
        public string Creator { get; set; }

        /// <summary>
        ///     The summary or abstract of the contents of the document. This might also be the comment section.
        /// </summary>
        public string Description { get; set; }

        internal string Identifier { get; set; }

        /// <summary>
        ///     A word or set of words describing the document.
        /// </summary>
        public string Keywords { get; set; }

        internal string Language { get; set; }

        /// <summary>
        ///     The document is last modified by this person.
        /// </summary>
        public string LastModifiedBy { get; set; }

        internal string LastPrinted { get; set; }

        internal string Modified { get; set; }

        internal string Revision { get; set; }

        /// <summary>
        ///     The topic of the contents of the document.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        ///     The title of the document.
        /// </summary>
        public string Title { get; set; }

        internal string Version { get; set; }

        internal void SetAllNull()
        {
            Category = string.Empty;
            ContentStatus = string.Empty;
            Created = string.Empty;
            Creator = string.Empty;
            Description = string.Empty;
            Identifier = string.Empty;
            Keywords = string.Empty;
            Language = string.Empty;
            LastModifiedBy = string.Empty;
            LastPrinted = string.Empty;
            Modified = string.Empty;
            Revision = string.Empty;
            Subject = string.Empty;
            Title = string.Empty;
            Version = string.Empty;
        }
    }
}