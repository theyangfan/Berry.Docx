using System;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    public class FootEndnoteFormat
    {
        private readonly Document _doc;
        private readonly Settings _settings;
        private readonly Section _section;
        private readonly NoteType _noteType;
        internal FootEndnoteFormat(Document doc, Settings settings, NoteType noteType)
        {
            _doc = doc;
            _settings = settings;
            _noteType = noteType;
        }

        internal FootEndnoteFormat(Document doc, Section section, NoteType noteType)
        {
            _doc = doc;
            _section = section;
            _noteType = noteType;
        }

        public FootEndnoteNumberRestartRule RestartRule
        {
            get
            {
                if(_noteType == NoteType.DocumentWideFootnote)
                {
                    W.FootnoteDocumentWideProperties fnPr =
                            _settings.XElement.Elements<W.FootnoteDocumentWideProperties>().FirstOrDefault();
                    if (fnPr?.NumberingRestart?.Val != null)
                        return Convert(fnPr.NumberingRestart.Val);
                }
                else if(_noteType == NoteType.DocumentWideEndnote)
                {
                    W.EndnoteDocumentWideProperties enPr =
                            _settings.XElement.Elements<W.EndnoteDocumentWideProperties>().FirstOrDefault();
                    if (enPr?.NumberingRestart?.Val != null)
                        return Convert(enPr.NumberingRestart.Val);
                }
                else if (_noteType == NoteType.SectionWideFootnote)
                {
                    W.FootnoteProperties fnPr = _section.XElement.Elements<W.FootnoteProperties>().FirstOrDefault();
                    if (fnPr?.NumberingRestart?.Val != null)
                    {
                        return Convert(fnPr.NumberingRestart.Val);
                    }
                    else
                    {
                        return _doc.FootnoteFormat.RestartRule;
                    }
                }
                else
                {
                    W.EndnoteProperties enPr = _section.XElement.Elements<W.EndnoteProperties>().FirstOrDefault();
                    if (enPr?.NumberingRestart?.Val != null)
                    {
                        return Convert(enPr.NumberingRestart.Val);
                    }
                    else
                    {
                        return _doc.EndnoteFormat.RestartRule;
                    }
                }
                return FootEndnoteNumberRestartRule.Continuous;
            }
            set
            {
                if (_noteType == NoteType.DocumentWideFootnote)
                {
                    W.FootnoteDocumentWideProperties fnPr =
                        _settings.XElement.Elements<W.FootnoteDocumentWideProperties>().FirstOrDefault();
                    if(fnPr == null)
                    {
                        fnPr = new W.FootnoteDocumentWideProperties();
                        _settings.XElement.AddChild(fnPr);
                    }
                    if(value == FootEndnoteNumberRestartRule.Continuous)
                        fnPr.Remove();
                    else
                        fnPr.NumberingRestart = new W.NumberingRestart() { Val = Convert(value) };
                    foreach (Section section in _doc.Sections)
                        section.FootnoteFormat.RestartRule = value;
                }
                else if (_noteType == NoteType.DocumentWideEndnote)
                {
                    W.EndnoteDocumentWideProperties enPr =
                        _settings.XElement.Elements<W.EndnoteDocumentWideProperties>().FirstOrDefault();
                    if (enPr == null)
                    {
                        enPr = new W.EndnoteDocumentWideProperties();
                        _settings.XElement.AddChild(enPr);
                    }
                    if (value == FootEndnoteNumberRestartRule.Continuous)
                        enPr.Remove();
                    else if (value == FootEndnoteNumberRestartRule.EachSection)
                        enPr.NumberingRestart = new W.NumberingRestart() { Val = Convert(value) };
                    foreach (Section section in _doc.Sections)
                        section.EndnoteFormat.RestartRule = value;
                }
                else if (_noteType == NoteType.SectionWideFootnote)
                {
                    W.FootnoteProperties fnPr = _section.XElement.Elements<W.FootnoteProperties>().FirstOrDefault();
                    if(fnPr == null)
                    {
                        fnPr = new W.FootnoteProperties();
                        _section.XElement.AddChild(fnPr);
                    }
                    if (value == FootEndnoteNumberRestartRule.Continuous
                            && _doc.FootnoteFormat.RestartRule == FootEndnoteNumberRestartRule.Continuous)
                        fnPr.Remove();
                    else
                        fnPr.NumberingRestart = new W.NumberingRestart() { Val = Convert(value) };
                }
                else
                {
                    W.EndnoteProperties enPr = _section.XElement.Elements<W.EndnoteProperties>().FirstOrDefault();
                    if (enPr == null)
                    {
                        enPr = new W.EndnoteProperties();
                        _section.XElement.AddChild(enPr);
                    }
                    if (value == FootEndnoteNumberRestartRule.Continuous
                            && _doc.EndnoteFormat.RestartRule == FootEndnoteNumberRestartRule.Continuous)
                        enPr.Remove();
                    else if(value == FootEndnoteNumberRestartRule.EachSection)
                        enPr.NumberingRestart = new W.NumberingRestart() { Val = Convert(value) };
                }
            }
        }

        private FootEndnoteNumberRestartRule Convert(W.RestartNumberValues val)
        {
            switch (val)
            {
                case W.RestartNumberValues.Continuous:
                    return FootEndnoteNumberRestartRule.Continuous;
                case W.RestartNumberValues.EachSection:
                    return FootEndnoteNumberRestartRule.EachSection;
                case W.RestartNumberValues.EachPage:
                    return FootEndnoteNumberRestartRule.EachPage;
                default:
                    return FootEndnoteNumberRestartRule.Continuous;
            }
        }

        private W.RestartNumberValues Convert(FootEndnoteNumberRestartRule rule)
        {
            switch (rule)
            {
                case FootEndnoteNumberRestartRule.Continuous:
                    return W.RestartNumberValues.Continuous;
                case FootEndnoteNumberRestartRule.EachSection:
                    return W.RestartNumberValues.EachSection;
                case FootEndnoteNumberRestartRule.EachPage:
                    return W.RestartNumberValues.EachPage;
                default:
                    return W.RestartNumberValues.Continuous;
            }
        }

    }

    internal enum NoteType
    {
        DocumentWideFootnote = 0,
        DocumentWideEndnote = 1,
        SectionWideFootnote = 2,
        SectionWideEndnote = 3
    }
}
