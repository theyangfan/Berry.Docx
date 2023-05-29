using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// This class specified the set of default paragraph and character format.
    /// </summary>
    public class DocDefaultFormat
    {
        private readonly CharacterFormat _cFormat;
        private readonly ParagraphFormat _pFormat;
        internal DocDefaultFormat(Document doc)
        {
            _cFormat = new CharacterFormat();
            _pFormat = new ParagraphFormat();

            W.DocDefaults defaults = doc.Package.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            if (defaults?.RunPropertiesDefault != null)
            {
                RunPropertiesHolder rHld = new RunPropertiesHolder(doc.Package, defaults.RunPropertiesDefault);

                if (rHld.FontNameAscii != null) _cFormat.FontNameAscii = rHld.FontNameAscii;
                if (rHld.FontNameEastAsia != null) _cFormat.FontNameEastAsia = rHld.FontNameEastAsia;
                if (rHld.FontNameHighAnsi != null) _cFormat.FontNameHighAnsi = rHld.FontNameHighAnsi;
                if (rHld.FontNameComplexScript != null) _cFormat.FontNameComplexScript = rHld.FontNameComplexScript;
                if (rHld.FontSize != null) _cFormat.FontSize = rHld.FontSize;
                if (rHld.FontSizeCs != null) _cFormat.FontSizeCs = rHld.FontSizeCs;
                if (rHld.Bold != null) _cFormat.Bold = rHld.Bold;
                if (rHld.BoldCs != null) _cFormat.BoldCs = rHld.BoldCs;
                if (rHld.Italic != null) _cFormat.Italic = rHld.Italic;
                if (rHld.ItalicCs != null) _cFormat.ItalicCs = rHld.ItalicCs;
                if (rHld.SubSuperScript != null) _cFormat.SubSuperScript = rHld.SubSuperScript;
                if (rHld.UnderlineStyle != null) _cFormat.UnderlineStyle = rHld.UnderlineStyle;
                if (rHld.TextColor != null) _cFormat.TextColor = rHld.TextColor;
                if (rHld.CharacterScale != null) _cFormat.CharacterScale = rHld.CharacterScale;
                if (rHld.CharacterSpacing != null) _cFormat.CharacterSpacing = rHld.CharacterSpacing;
                if (rHld.Position != null) _cFormat.Position = rHld.Position;
                if (rHld.IsHidden != null) _cFormat.IsHidden = rHld.IsHidden;
                if (rHld.SnapToGrid != null) _cFormat.SnapToGrid = rHld.SnapToGrid;

                _cFormat.Border = new Border(doc, defaults.RunPropertiesDefault);
            }
            if (defaults?.ParagraphPropertiesDefault != null)
            {

            }
        }

        /// <summary>
        /// The default character format.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// The default paragraph format.
        /// </summary>
        public ParagraphFormat ParagraphFormat => _pFormat;
    }
}
