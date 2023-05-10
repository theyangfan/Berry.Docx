using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Hyperlink : ParagraphItem
    {
        private readonly W.Hyperlink _hyperlink;
        public Hyperlink(Document doc, W.Hyperlink hyperlink) : base(doc, hyperlink)
        {
            _hyperlink = hyperlink;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Hyperlink;

        public HyperlinkTargetType TargetType
        {
            get
            {
                if(_hyperlink.Id != null)
                {
                    return HyperlinkTargetType.Hyperlink;
                }
                else if(_hyperlink.Anchor != null)
                {
                    return HyperlinkTargetType.Bookmark;
                }
                return HyperlinkTargetType.Invalid;
            }
        }

        public string Target
        {
            get
            {
                return "";
            }
        }

        public bool AddToViewedHistory
        {
            get
            {
                return false;
            }
        }
    }
}
