using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class DocDefaultFormat
    {
        private readonly CharacterFormat _cFormat;
        private readonly ParagraphFormat _pFormat;
        internal DocDefaultFormat(Document doc)
        {
            _cFormat = new CharacterFormat();
            _pFormat = new ParagraphFormat();

            W.DocDefaults defaults = doc.Package.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            if(defaults?.RunPropertiesDefault != null)
            {
                RunPropertiesHolder rHld = new RunPropertiesHolder(doc.Package, defaults.RunPropertiesDefault);
                if(rHld.FontNameEastAsia != null)
                    _cFormat.FontNameEastAsia = rHld.FontNameEastAsia;
                if(rHld.FontNameAscii != null)
                    _cFormat.FontNameAscii = rHld.FontNameAscii;
                if(rHld.FontSize != null)
                    _cFormat.FontSize = rHld.FontSize;
                if (rHld.FontSizeCs != null)
                    _cFormat.FontSizeCs = rHld.FontSizeCs;
                if (rHld.Bold != null)
                    _cFormat.Bold = rHld.Bold;
                if (rHld.Italic != null)
                    _cFormat.Italic = rHld.Italic;
                if(rHld.SubSuperScript != null)
                    _cFormat.SubSuperScript = rHld.SubSuperScript;
                if (rHld.UnderlineStyle != null)
                    _cFormat.UnderlineStyle = rHld.UnderlineStyle;
                if(!rHld.TextColor.IsEmpty)
                    _cFormat.TextColor = rHld.TextColor;
                if(rHld.AutoTextColor != null)
                    _cFormat.AutoTextColor = rHld.AutoTextColor;
                if (rHld.CharacterScale != null)
                    _cFormat.CharacterScale = rHld.CharacterScale;
                if (rHld.CharacterSpacing != null)
                    _cFormat.CharacterSpacing = rHld.CharacterSpacing;
                if (rHld.Position != null)
                    _cFormat.Position = rHld.Position;
            }
            if(defaults?.ParagraphPropertiesDefault != null)
            {

            }
        }

        public CharacterFormat CharacterFormat => _cFormat;

        public ParagraphFormat ParagraphFormat => _pFormat;
    }
}
