using System.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.misc
{
    /// <summary>
    ///     Simple settings for themes.
    /// </summary>
    public class SLThemeSettings
    {
        /// <summary>
        ///     Initialize an instance of SLThemeSettings.
        /// </summary>
        public SLThemeSettings()
        {
            SetTheme(SLThemeTypeValues.Office);
            ThemeName = "SpreadsheetLight Custom";
        }

        /// <summary>
        ///     Initialize an instance of SLThemeSettings with a given theme.
        /// </summary>
        /// <param name="ThemeType">A built-in theme.</param>
        public SLThemeSettings(SLThemeTypeValues ThemeType)
        {
            SetTheme(ThemeType);
        }

        /// <summary>
        ///     The theme name.
        /// </summary>
        public string ThemeName { get; set; }

        /// <summary>
        ///     The major latin font.
        /// </summary>
        public string MajorLatinFont { get; set; }

        /// <summary>
        ///     The minor latin font.
        /// </summary>
        public string MinorLatinFont { get; set; }

        /// <summary>
        ///     Typically pure black.
        /// </summary>
        public Color Dark1Color { get; set; }

        /// <summary>
        ///     Typically pure white.
        /// </summary>
        public Color Light1Color { get; set; }

        /// <summary>
        ///     A dark color that still has visual contrast against light tints of the accent colors.
        /// </summary>
        public Color Dark2Color { get; set; }

        /// <summary>
        ///     A light color that still has visual contrast against dark tints of the accent colors.
        /// </summary>
        public Color Light2Color { get; set; }

        /// <summary>
        ///     Accent1 color.
        /// </summary>
        public Color Accent1Color { get; set; }

        /// <summary>
        ///     Accent2 color.
        /// </summary>
        public Color Accent2Color { get; set; }

        /// <summary>
        ///     Accent3 color.
        /// </summary>
        public Color Accent3Color { get; set; }

        /// <summary>
        ///     Accent4 color.
        /// </summary>
        public Color Accent4Color { get; set; }

        /// <summary>
        ///     Accent5 color.
        /// </summary>
        public Color Accent5Color { get; set; }

        /// <summary>
        ///     Accent6 color.
        /// </summary>
        public Color Accent6Color { get; set; }

        /// <summary>
        ///     Color of a hyperlink.
        /// </summary>
        public Color Hyperlink { get; set; }

        /// <summary>
        ///     Color of a followed hyperlink.
        /// </summary>
        public Color FollowedHyperlinkColor { get; set; }

        private void SetTheme(SLThemeTypeValues ThemeType)
        {
            switch (ThemeType)
            {
                case SLThemeTypeValues.Office:
                    ThemeName = SLConstants.OfficeThemeName;
                    MajorLatinFont = SLConstants.OfficeThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.OfficeThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.OfficeThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.OfficeThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.OfficeThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.OfficeThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.OfficeThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.OfficeThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.OfficeThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.OfficeThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.OfficeThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.OfficeThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.OfficeThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.OfficeThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Office2013:
                    ThemeName = SLConstants.Office2013ThemeName;
                    MajorLatinFont = SLConstants.Office2013ThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.Office2013ThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.Office2013ThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.Office2013ThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.Office2013ThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.Office2013ThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.Office2013ThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.Office2013ThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.Office2013ThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.Office2013ThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.Office2013ThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.Office2013ThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.Office2013ThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.Office2013ThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Adjacency:
                    ThemeName = SLConstants.AdjacencyThemeName;
                    MajorLatinFont = SLConstants.AdjacencyThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.AdjacencyThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.AdjacencyThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.AdjacencyThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.AdjacencyThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.AdjacencyThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.AdjacencyThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.AdjacencyThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.AdjacencyThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.AdjacencyThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.AdjacencyThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.AdjacencyThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.AdjacencyThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.AdjacencyThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Angles:
                    ThemeName = SLConstants.AnglesThemeName;
                    MajorLatinFont = SLConstants.AnglesThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.AnglesThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.AnglesThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.AnglesThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.AnglesThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.AnglesThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.AnglesThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.AnglesThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.AnglesThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.AnglesThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.AnglesThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.AnglesThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.AnglesThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.AnglesThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Apex:
                    ThemeName = SLConstants.ApexThemeName;
                    MajorLatinFont = SLConstants.ApexThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ApexThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ApexThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ApexThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ApexThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ApexThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ApexThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ApexThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ApexThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ApexThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ApexThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ApexThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ApexThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ApexThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Apothecary:
                    ThemeName = SLConstants.ApothecaryThemeName;
                    MajorLatinFont = SLConstants.ApothecaryThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ApothecaryThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ApothecaryThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ApothecaryThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ApothecaryThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ApothecaryThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ApothecaryThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ApothecaryThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ApothecaryThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ApothecaryThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ApothecaryThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ApothecaryThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ApothecaryThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ApothecaryThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Aspect:
                    ThemeName = SLConstants.AspectThemeName;
                    MajorLatinFont = SLConstants.AspectThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.AspectThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.AspectThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.AspectThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.AspectThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.AspectThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.AspectThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.AspectThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.AspectThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.AspectThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.AspectThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.AspectThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.AspectThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.AspectThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Austin:
                    ThemeName = SLConstants.AustinThemeName;
                    MajorLatinFont = SLConstants.AustinThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.AustinThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.AustinThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.AustinThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.AustinThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.AustinThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.AustinThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.AustinThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.AustinThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.AustinThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.AustinThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.AustinThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.AustinThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.AustinThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.BlackTie:
                    ThemeName = SLConstants.BlackTieThemeName;
                    MajorLatinFont = SLConstants.BlackTieThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.BlackTieThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.BlackTieThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.BlackTieThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.BlackTieThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.BlackTieThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.BlackTieThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.BlackTieThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.BlackTieThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.BlackTieThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.BlackTieThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.BlackTieThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.BlackTieThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.BlackTieThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Civic:
                    ThemeName = SLConstants.CivicThemeName;
                    MajorLatinFont = SLConstants.CivicThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.CivicThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.CivicThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.CivicThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.CivicThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.CivicThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.CivicThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.CivicThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.CivicThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.CivicThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.CivicThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.CivicThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.CivicThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.CivicThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Clarity:
                    ThemeName = SLConstants.ClarityThemeName;
                    MajorLatinFont = SLConstants.ClarityThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ClarityThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ClarityThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ClarityThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ClarityThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ClarityThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ClarityThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ClarityThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ClarityThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ClarityThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ClarityThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ClarityThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ClarityThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ClarityThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Composite:
                    ThemeName = SLConstants.CompositeThemeName;
                    MajorLatinFont = SLConstants.CompositeThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.CompositeThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.CompositeThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.CompositeThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.CompositeThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.CompositeThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.CompositeThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.CompositeThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.CompositeThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.CompositeThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.CompositeThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.CompositeThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.CompositeThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.CompositeThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Concourse:
                    ThemeName = SLConstants.ConcourseThemeName;
                    MajorLatinFont = SLConstants.ConcourseThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ConcourseThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ConcourseThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ConcourseThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ConcourseThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ConcourseThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ConcourseThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ConcourseThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ConcourseThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ConcourseThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ConcourseThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ConcourseThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ConcourseThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ConcourseThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Couture:
                    ThemeName = SLConstants.CoutureThemeName;
                    MajorLatinFont = SLConstants.CoutureThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.CoutureThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.CoutureThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.CoutureThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.CoutureThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.CoutureThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.CoutureThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.CoutureThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.CoutureThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.CoutureThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.CoutureThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.CoutureThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.CoutureThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.CoutureThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Elemental:
                    ThemeName = SLConstants.ElementalThemeName;
                    MajorLatinFont = SLConstants.ElementalThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ElementalThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ElementalThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ElementalThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ElementalThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ElementalThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ElementalThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ElementalThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ElementalThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ElementalThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ElementalThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ElementalThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ElementalThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ElementalThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Equity:
                    ThemeName = SLConstants.EquityThemeName;
                    MajorLatinFont = SLConstants.EquityThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.EquityThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.EquityThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.EquityThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.EquityThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.EquityThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.EquityThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.EquityThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.EquityThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.EquityThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.EquityThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.EquityThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.EquityThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.EquityThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Essential:
                    ThemeName = SLConstants.EssentialThemeName;
                    MajorLatinFont = SLConstants.EssentialThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.EssentialThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.EssentialThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.EssentialThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.EssentialThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.EssentialThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.EssentialThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.EssentialThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.EssentialThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.EssentialThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.EssentialThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.EssentialThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.EssentialThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.EssentialThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Executive:
                    ThemeName = SLConstants.ExecutiveThemeName;
                    MajorLatinFont = SLConstants.ExecutiveThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ExecutiveThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ExecutiveThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ExecutiveThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ExecutiveThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ExecutiveThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ExecutiveThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ExecutiveThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ExecutiveThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ExecutiveThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ExecutiveThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ExecutiveThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ExecutiveThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ExecutiveThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Facet:
                    ThemeName = SLConstants.FacetThemeName;
                    MajorLatinFont = SLConstants.FacetThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.FacetThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.FacetThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.FacetThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.FacetThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.FacetThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.FacetThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.FacetThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.FacetThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.FacetThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.FacetThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.FacetThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.FacetThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.FacetThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Flow:
                    ThemeName = SLConstants.FlowThemeName;
                    MajorLatinFont = SLConstants.FlowThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.FlowThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.FlowThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.FlowThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.FlowThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.FlowThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.FlowThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.FlowThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.FlowThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.FlowThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.FlowThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.FlowThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.FlowThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.FlowThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Foundry:
                    ThemeName = SLConstants.FoundryThemeName;
                    MajorLatinFont = SLConstants.FoundryThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.FoundryThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.FoundryThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.FoundryThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.FoundryThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.FoundryThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.FoundryThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.FoundryThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.FoundryThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.FoundryThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.FoundryThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.FoundryThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.FoundryThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.FoundryThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Grid:
                    ThemeName = SLConstants.GridThemeName;
                    MajorLatinFont = SLConstants.GridThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.GridThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.GridThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.GridThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.GridThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.GridThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.GridThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.GridThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.GridThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.GridThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.GridThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.GridThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.GridThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.GridThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Hardcover:
                    ThemeName = SLConstants.HardcoverThemeName;
                    MajorLatinFont = SLConstants.HardcoverThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.HardcoverThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.HardcoverThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.HardcoverThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.HardcoverThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.HardcoverThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.HardcoverThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.HardcoverThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.HardcoverThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.HardcoverThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.HardcoverThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.HardcoverThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.HardcoverThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.HardcoverThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Horizon:
                    ThemeName = SLConstants.HorizonThemeName;
                    MajorLatinFont = SLConstants.HorizonThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.HorizonThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.HorizonThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.HorizonThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.HorizonThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.HorizonThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.HorizonThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.HorizonThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.HorizonThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.HorizonThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.HorizonThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.HorizonThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.HorizonThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.HorizonThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Integral:
                    ThemeName = SLConstants.IntegralThemeName;
                    MajorLatinFont = SLConstants.IntegralThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.IntegralThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.IntegralThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.IntegralThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.IntegralThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.IntegralThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.IntegralThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.IntegralThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.IntegralThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.IntegralThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.IntegralThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.IntegralThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.IntegralThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.IntegralThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Ion:
                    ThemeName = SLConstants.IonThemeName;
                    MajorLatinFont = SLConstants.IonThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.IonThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.IonThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.IonThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.IonThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.IonThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.IonThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.IonThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.IonThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.IonThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.IonThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.IonThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.IonThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.IonThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.IonBoardroom:
                    ThemeName = SLConstants.IonBoardroomThemeName;
                    MajorLatinFont = SLConstants.IonBoardroomThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.IonBoardroomThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.IonBoardroomThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.IonBoardroomThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.IonBoardroomThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.IonBoardroomThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.IonBoardroomThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.IonBoardroomThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.IonBoardroomThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.IonBoardroomThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.IonBoardroomThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.IonBoardroomThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.IonBoardroomThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.IonBoardroomThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Median:
                    ThemeName = SLConstants.MedianThemeName;
                    MajorLatinFont = SLConstants.MedianThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MedianThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MedianThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MedianThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MedianThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MedianThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MedianThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MedianThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MedianThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MedianThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MedianThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MedianThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MedianThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MedianThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Metro:
                    ThemeName = SLConstants.MetroThemeName;
                    MajorLatinFont = SLConstants.MetroThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MetroThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MetroThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MetroThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MetroThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MetroThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MetroThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MetroThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MetroThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MetroThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MetroThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MetroThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MetroThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MetroThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Module:
                    ThemeName = SLConstants.ModuleThemeName;
                    MajorLatinFont = SLConstants.ModuleThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ModuleThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ModuleThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ModuleThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ModuleThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ModuleThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ModuleThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ModuleThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ModuleThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ModuleThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ModuleThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ModuleThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ModuleThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ModuleThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Newsprint:
                    ThemeName = SLConstants.NewsprintThemeName;
                    MajorLatinFont = SLConstants.NewsprintThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.NewsprintThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.NewsprintThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.NewsprintThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.NewsprintThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.NewsprintThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.NewsprintThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.NewsprintThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.NewsprintThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.NewsprintThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.NewsprintThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.NewsprintThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.NewsprintThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.NewsprintThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Opulent:
                    ThemeName = SLConstants.OpulentThemeName;
                    MajorLatinFont = SLConstants.OpulentThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.OpulentThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.OpulentThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.OpulentThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.OpulentThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.OpulentThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.OpulentThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.OpulentThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.OpulentThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.OpulentThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.OpulentThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.OpulentThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.OpulentThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.OpulentThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Organic:
                    ThemeName = SLConstants.OrganicThemeName;
                    MajorLatinFont = SLConstants.OrganicThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.OrganicThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.OrganicThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.OrganicThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.OrganicThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.OrganicThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.OrganicThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.OrganicThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.OrganicThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.OrganicThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.OrganicThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.OrganicThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.OrganicThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.OrganicThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Oriel:
                    ThemeName = SLConstants.OrielThemeName;
                    MajorLatinFont = SLConstants.OrielThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.OrielThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.OrielThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.OrielThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.OrielThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.OrielThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.OrielThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.OrielThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.OrielThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.OrielThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.OrielThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.OrielThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.OrielThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.OrielThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Origin:
                    ThemeName = SLConstants.OriginThemeName;
                    MajorLatinFont = SLConstants.OriginThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.OriginThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.OriginThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.OriginThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.OriginThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.OriginThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.OriginThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.OriginThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.OriginThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.OriginThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.OriginThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.OriginThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.OriginThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.OriginThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Paper:
                    ThemeName = SLConstants.PaperThemeName;
                    MajorLatinFont = SLConstants.PaperThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.PaperThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.PaperThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.PaperThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.PaperThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.PaperThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.PaperThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.PaperThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.PaperThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.PaperThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.PaperThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.PaperThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.PaperThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.PaperThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Perspective:
                    ThemeName = SLConstants.PerspectiveThemeName;
                    MajorLatinFont = SLConstants.PerspectiveThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.PerspectiveThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.PerspectiveThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.PerspectiveThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.PerspectiveThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.PerspectiveThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.PerspectiveThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.PerspectiveThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.PerspectiveThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.PerspectiveThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.PerspectiveThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.PerspectiveThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.PerspectiveThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.PerspectiveThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Pushpin:
                    ThemeName = SLConstants.PushpinThemeName;
                    MajorLatinFont = SLConstants.PushpinThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.PushpinThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.PushpinThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.PushpinThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.PushpinThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.PushpinThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.PushpinThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.PushpinThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.PushpinThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.PushpinThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.PushpinThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.PushpinThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.PushpinThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.PushpinThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Retrospect:
                    ThemeName = SLConstants.RetrospectThemeName;
                    MajorLatinFont = SLConstants.RetrospectThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.RetrospectThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.RetrospectThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.RetrospectThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.RetrospectThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.RetrospectThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.RetrospectThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.RetrospectThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.RetrospectThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.RetrospectThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.RetrospectThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.RetrospectThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.RetrospectThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.RetrospectThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Slice:
                    ThemeName = SLConstants.SliceThemeName;
                    MajorLatinFont = SLConstants.SliceThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SliceThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SliceThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SliceThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SliceThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SliceThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SliceThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SliceThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SliceThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SliceThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SliceThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SliceThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SliceThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SliceThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Slipstream:
                    ThemeName = SLConstants.SlipstreamThemeName;
                    MajorLatinFont = SLConstants.SlipstreamThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SlipstreamThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SlipstreamThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SlipstreamThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SlipstreamThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SlipstreamThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SlipstreamThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SlipstreamThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SlipstreamThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SlipstreamThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SlipstreamThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SlipstreamThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SlipstreamThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SlipstreamThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Solstice:
                    ThemeName = SLConstants.SolsticeThemeName;
                    MajorLatinFont = SLConstants.SolsticeThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SolsticeThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SolsticeThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SolsticeThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SolsticeThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SolsticeThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SolsticeThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SolsticeThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SolsticeThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SolsticeThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SolsticeThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SolsticeThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SolsticeThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SolsticeThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Technic:
                    ThemeName = SLConstants.TechnicThemeName;
                    MajorLatinFont = SLConstants.TechnicThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.TechnicThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.TechnicThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.TechnicThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.TechnicThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.TechnicThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.TechnicThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.TechnicThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.TechnicThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.TechnicThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.TechnicThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.TechnicThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.TechnicThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.TechnicThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Thatch:
                    ThemeName = SLConstants.ThatchThemeName;
                    MajorLatinFont = SLConstants.ThatchThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ThatchThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ThatchThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ThatchThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ThatchThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ThatchThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ThatchThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ThatchThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ThatchThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ThatchThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ThatchThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ThatchThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ThatchThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ThatchThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Trek:
                    ThemeName = SLConstants.TrekThemeName;
                    MajorLatinFont = SLConstants.TrekThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.TrekThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.TrekThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.TrekThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.TrekThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.TrekThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.TrekThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.TrekThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.TrekThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.TrekThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.TrekThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.TrekThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.TrekThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.TrekThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Urban:
                    ThemeName = SLConstants.UrbanThemeName;
                    MajorLatinFont = SLConstants.UrbanThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.UrbanThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.UrbanThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.UrbanThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.UrbanThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.UrbanThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.UrbanThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.UrbanThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.UrbanThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.UrbanThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.UrbanThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.UrbanThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.UrbanThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.UrbanThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Verve:
                    ThemeName = SLConstants.VerveThemeName;
                    MajorLatinFont = SLConstants.VerveThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.VerveThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.VerveThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.VerveThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.VerveThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.VerveThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.VerveThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.VerveThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.VerveThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.VerveThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.VerveThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.VerveThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.VerveThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.VerveThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Waveform:
                    ThemeName = SLConstants.WaveformThemeName;
                    MajorLatinFont = SLConstants.WaveformThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.WaveformThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.WaveformThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.WaveformThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.WaveformThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.WaveformThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.WaveformThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.WaveformThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.WaveformThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.WaveformThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.WaveformThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.WaveformThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.WaveformThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.WaveformThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Wisp:
                    ThemeName = SLConstants.WispThemeName;
                    MajorLatinFont = SLConstants.WispThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.WispThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.WispThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.WispThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.WispThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.WispThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.WispThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.WispThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.WispThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.WispThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.WispThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.WispThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.WispThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.WispThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Autumn:
                    ThemeName = SLConstants.AutumnThemeName;
                    MajorLatinFont = SLConstants.AutumnThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.AutumnThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.AutumnThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.AutumnThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.AutumnThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.AutumnThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.AutumnThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.AutumnThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.AutumnThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.AutumnThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.AutumnThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.AutumnThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.AutumnThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.AutumnThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Banded:
                    ThemeName = SLConstants.BandedThemeName;
                    MajorLatinFont = SLConstants.BandedThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.BandedThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.BandedThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.BandedThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.BandedThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.BandedThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.BandedThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.BandedThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.BandedThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.BandedThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.BandedThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.BandedThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.BandedThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.BandedThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Basis:
                    ThemeName = SLConstants.BasisThemeName;
                    MajorLatinFont = SLConstants.BasisThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.BasisThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.BasisThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.BasisThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.BasisThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.BasisThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.BasisThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.BasisThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.BasisThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.BasisThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.BasisThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.BasisThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.BasisThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.BasisThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Berlin:
                    ThemeName = SLConstants.BerlinThemeName;
                    MajorLatinFont = SLConstants.BerlinThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.BerlinThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.BerlinThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.BerlinThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.BerlinThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.BerlinThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.BerlinThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.BerlinThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.BerlinThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.BerlinThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.BerlinThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.BerlinThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.BerlinThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.BerlinThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Celestial:
                    ThemeName = SLConstants.CelestialThemeName;
                    MajorLatinFont = SLConstants.CelestialThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.CelestialThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.CelestialThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.CelestialThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.CelestialThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.CelestialThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.CelestialThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.CelestialThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.CelestialThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.CelestialThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.CelestialThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.CelestialThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.CelestialThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.CelestialThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Circuit:
                    ThemeName = SLConstants.CircuitThemeName;
                    MajorLatinFont = SLConstants.CircuitThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.CircuitThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.CircuitThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.CircuitThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.CircuitThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.CircuitThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.CircuitThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.CircuitThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.CircuitThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.CircuitThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.CircuitThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.CircuitThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.CircuitThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.CircuitThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Damask:
                    ThemeName = SLConstants.DamaskThemeName;
                    MajorLatinFont = SLConstants.DamaskThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.DamaskThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.DamaskThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.DamaskThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.DamaskThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.DamaskThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.DamaskThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.DamaskThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.DamaskThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.DamaskThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.DamaskThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.DamaskThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.DamaskThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.DamaskThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Decatur:
                    ThemeName = SLConstants.DecaturThemeName;
                    MajorLatinFont = SLConstants.DecaturThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.DecaturThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.DecaturThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.DecaturThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.DecaturThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.DecaturThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.DecaturThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.DecaturThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.DecaturThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.DecaturThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.DecaturThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.DecaturThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.DecaturThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.DecaturThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Depth:
                    ThemeName = SLConstants.DepthThemeName;
                    MajorLatinFont = SLConstants.DepthThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.DepthThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.DepthThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.DepthThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.DepthThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.DepthThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.DepthThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.DepthThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.DepthThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.DepthThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.DepthThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.DepthThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.DepthThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.DepthThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Dividend:
                    ThemeName = SLConstants.DividendThemeName;
                    MajorLatinFont = SLConstants.DividendThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.DividendThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.DividendThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.DividendThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.DividendThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.DividendThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.DividendThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.DividendThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.DividendThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.DividendThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.DividendThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.DividendThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.DividendThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.DividendThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Droplet:
                    ThemeName = SLConstants.DropletThemeName;
                    MajorLatinFont = SLConstants.DropletThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.DropletThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.DropletThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.DropletThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.DropletThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.DropletThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.DropletThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.DropletThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.DropletThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.DropletThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.DropletThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.DropletThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.DropletThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.DropletThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Frame:
                    ThemeName = SLConstants.FrameThemeName;
                    MajorLatinFont = SLConstants.FrameThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.FrameThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.FrameThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.FrameThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.FrameThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.FrameThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.FrameThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.FrameThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.FrameThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.FrameThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.FrameThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.FrameThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.FrameThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.FrameThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Kilter:
                    ThemeName = SLConstants.KilterThemeName;
                    MajorLatinFont = SLConstants.KilterThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.KilterThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.KilterThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.KilterThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.KilterThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.KilterThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.KilterThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.KilterThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.KilterThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.KilterThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.KilterThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.KilterThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.KilterThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.KilterThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Macro:
                    ThemeName = SLConstants.MacroThemeName;
                    MajorLatinFont = SLConstants.MacroThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MacroThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MacroThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MacroThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MacroThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MacroThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MacroThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MacroThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MacroThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MacroThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MacroThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MacroThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MacroThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MacroThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.MainEvent:
                    ThemeName = SLConstants.MainEventThemeName;
                    MajorLatinFont = SLConstants.MainEventThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MainEventThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MainEventThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MainEventThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MainEventThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MainEventThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MainEventThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MainEventThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MainEventThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MainEventThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MainEventThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MainEventThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MainEventThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MainEventThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Mesh:
                    ThemeName = SLConstants.MeshThemeName;
                    MajorLatinFont = SLConstants.MeshThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MeshThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MeshThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MeshThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MeshThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MeshThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MeshThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MeshThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MeshThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MeshThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MeshThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MeshThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MeshThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MeshThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Metropolitan:
                    ThemeName = SLConstants.MetropolitanThemeName;
                    MajorLatinFont = SLConstants.MetropolitanThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MetropolitanThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MetropolitanThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MetropolitanThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MetropolitanThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MetropolitanThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MetropolitanThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MetropolitanThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MetropolitanThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MetropolitanThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MetropolitanThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MetropolitanThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MetropolitanThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MetropolitanThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Mylar:
                    ThemeName = SLConstants.MylarThemeName;
                    MajorLatinFont = SLConstants.MylarThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.MylarThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.MylarThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.MylarThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.MylarThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.MylarThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.MylarThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.MylarThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.MylarThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.MylarThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.MylarThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.MylarThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.MylarThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.MylarThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Parallax:
                    ThemeName = SLConstants.ParallaxThemeName;
                    MajorLatinFont = SLConstants.ParallaxThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ParallaxThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ParallaxThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ParallaxThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ParallaxThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ParallaxThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ParallaxThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ParallaxThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ParallaxThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ParallaxThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ParallaxThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ParallaxThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ParallaxThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ParallaxThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Quotable:
                    ThemeName = SLConstants.QuotableThemeName;
                    MajorLatinFont = SLConstants.QuotableThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.QuotableThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.QuotableThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.QuotableThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.QuotableThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.QuotableThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.QuotableThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.QuotableThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.QuotableThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.QuotableThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.QuotableThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.QuotableThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.QuotableThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.QuotableThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Savon:
                    ThemeName = SLConstants.SavonThemeName;
                    MajorLatinFont = SLConstants.SavonThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SavonThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SavonThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SavonThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SavonThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SavonThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SavonThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SavonThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SavonThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SavonThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SavonThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SavonThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SavonThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SavonThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Sketchbook:
                    ThemeName = SLConstants.SketchbookThemeName;
                    MajorLatinFont = SLConstants.SketchbookThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SketchbookThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SketchbookThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SketchbookThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SketchbookThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SketchbookThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SketchbookThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SketchbookThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SketchbookThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SketchbookThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SketchbookThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SketchbookThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SketchbookThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SketchbookThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Slate:
                    ThemeName = SLConstants.SlateThemeName;
                    MajorLatinFont = SLConstants.SlateThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SlateThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SlateThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SlateThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SlateThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SlateThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SlateThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SlateThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SlateThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SlateThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SlateThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SlateThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SlateThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SlateThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Soho:
                    ThemeName = SLConstants.SohoThemeName;
                    MajorLatinFont = SLConstants.SohoThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SohoThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SohoThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SohoThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SohoThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SohoThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SohoThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SohoThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SohoThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SohoThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SohoThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SohoThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SohoThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SohoThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Spring:
                    ThemeName = SLConstants.SpringThemeName;
                    MajorLatinFont = SLConstants.SpringThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SpringThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SpringThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SpringThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SpringThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SpringThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SpringThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SpringThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SpringThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SpringThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SpringThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SpringThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SpringThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SpringThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Summer:
                    ThemeName = SLConstants.SummerThemeName;
                    MajorLatinFont = SLConstants.SummerThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.SummerThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.SummerThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.SummerThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.SummerThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.SummerThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.SummerThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.SummerThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.SummerThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.SummerThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.SummerThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.SummerThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.SummerThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.SummerThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Thermal:
                    ThemeName = SLConstants.ThermalThemeName;
                    MajorLatinFont = SLConstants.ThermalThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ThermalThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ThermalThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ThermalThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ThermalThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ThermalThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ThermalThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ThermalThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ThermalThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ThermalThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ThermalThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ThermalThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ThermalThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ThermalThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Tradeshow:
                    ThemeName = SLConstants.TradeshowThemeName;
                    MajorLatinFont = SLConstants.TradeshowThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.TradeshowThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.TradeshowThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.TradeshowThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.TradeshowThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.TradeshowThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.TradeshowThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.TradeshowThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.TradeshowThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.TradeshowThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.TradeshowThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.TradeshowThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.TradeshowThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.TradeshowThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.UrbanPop:
                    ThemeName = SLConstants.UrbanPopThemeName;
                    MajorLatinFont = SLConstants.UrbanPopThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.UrbanPopThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.UrbanPopThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.UrbanPopThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.UrbanPopThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.UrbanPopThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.UrbanPopThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.UrbanPopThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.UrbanPopThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.UrbanPopThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.UrbanPopThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.UrbanPopThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.UrbanPopThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.UrbanPopThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.VaporTrail:
                    ThemeName = SLConstants.VaporTrailThemeName;
                    MajorLatinFont = SLConstants.VaporTrailThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.VaporTrailThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.VaporTrailThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.VaporTrailThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.VaporTrailThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.VaporTrailThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.VaporTrailThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.VaporTrailThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.VaporTrailThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.VaporTrailThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.VaporTrailThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.VaporTrailThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.VaporTrailThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.VaporTrailThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.View:
                    ThemeName = SLConstants.ViewThemeName;
                    MajorLatinFont = SLConstants.ViewThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.ViewThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.ViewThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.ViewThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.ViewThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.ViewThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.ViewThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.ViewThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.ViewThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.ViewThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.ViewThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.ViewThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.ViewThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.ViewThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.Winter:
                    ThemeName = SLConstants.WinterThemeName;
                    MajorLatinFont = SLConstants.WinterThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.WinterThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.WinterThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.WinterThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.WinterThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.WinterThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.WinterThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.WinterThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.WinterThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.WinterThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.WinterThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.WinterThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.WinterThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.WinterThemeFollowedHyperlinkColor);
                    break;
                case SLThemeTypeValues.WoodType:
                    ThemeName = SLConstants.WoodTypeThemeName;
                    MajorLatinFont = SLConstants.WoodTypeThemeMajorLatinFont;
                    MinorLatinFont = SLConstants.WoodTypeThemeMinorLatinFont;
                    Dark1Color = SLTool.ToColor(SLConstants.WoodTypeThemeDark1Color);
                    Light1Color = SLTool.ToColor(SLConstants.WoodTypeThemeLight1Color);
                    Dark2Color = SLTool.ToColor(SLConstants.WoodTypeThemeDark2Color);
                    Light2Color = SLTool.ToColor(SLConstants.WoodTypeThemeLight2Color);
                    Accent1Color = SLTool.ToColor(SLConstants.WoodTypeThemeAccent1Color);
                    Accent2Color = SLTool.ToColor(SLConstants.WoodTypeThemeAccent2Color);
                    Accent3Color = SLTool.ToColor(SLConstants.WoodTypeThemeAccent3Color);
                    Accent4Color = SLTool.ToColor(SLConstants.WoodTypeThemeAccent4Color);
                    Accent5Color = SLTool.ToColor(SLConstants.WoodTypeThemeAccent5Color);
                    Accent6Color = SLTool.ToColor(SLConstants.WoodTypeThemeAccent6Color);
                    Hyperlink = SLTool.ToColor(SLConstants.WoodTypeThemeHyperlink);
                    FollowedHyperlinkColor = SLTool.ToColor(SLConstants.WoodTypeThemeFollowedHyperlinkColor);
                    break;
            }
        }
    }
}