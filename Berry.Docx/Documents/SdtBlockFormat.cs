using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class SdtBlockFormat
    {
        private readonly DocPartObject _docPartObject;
        internal SdtBlockFormat(Document doc, W.SdtProperties sdtPr)
        {
            _docPartObject = new DocPartObject(doc, sdtPr);
        }

        public DocPartObject DocPart => _docPartObject;
    }

    public class DocPartObject
    {
        private readonly W.SdtProperties _sdtPr;
        public DocPartObject(Document doc, W.SdtProperties sdtPr)
        {
            _sdtPr = sdtPr;
        }

        public string GalleryFilter
        {
            get
            {
                var gallery = _sdtPr.GetFirstChild<W.SdtContentDocPartObject>()?.GetFirstChild<W.DocPartGallery>();
                if(gallery == null) return null;
                return gallery.Val;
            }
            set
            {
                var docPartObj = _sdtPr.GetFirstChild<W.SdtContentDocPartObject>();
                if(docPartObj == null)
                {
                    docPartObj = new W.SdtContentDocPartObject();
                    _sdtPr.AddChild(docPartObj);
                }
                var gallery = docPartObj.GetFirstChild<W.DocPartGallery>();
                if(gallery == null)
                {
                    gallery = new W.DocPartGallery();
                    docPartObj.AddChild(gallery);
                }
                gallery.Val = value;
            }
        }

    }
}
