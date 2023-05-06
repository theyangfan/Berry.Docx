using System;
using System.Linq;
using System.Collections.Generic;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wpc = DocumentFormat.OpenXml.Office2010.Word.DrawingCanvas;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;
using Berry.Docx.Collections;
using Berry.Docx.Documents;
using Berry.Docx.Field;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a base class that all document item objects derive from.
    /// </summary>
    public abstract class DocumentItem : DocumentObject
    {
        #region Private Members
        private readonly Document _doc;
        private readonly O.OpenXmlElement _element;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the DocumentItem class using the supplied underlying OpenXmlElement.
        /// </summary>
        /// <param name="ownerDoc">Owner document</param>
        /// <param name="ele">Underlying OpenXmlElement</param>
        public DocumentItem(Document ownerDoc, O.OpenXmlElement ele)
            : base(ownerDoc, ele)
        {
            _doc = ownerDoc;
            _element = ele;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets all the child objects of the current item.
        /// </summary>
        public override DocumentObjectCollection ChildObjects => new DocumentItemCollection(_element, ChildItems());

        /// <summary>
        /// Gets the object that immediately precedes the current object. 
        /// Returns null if there is no preceding object.
        /// </summary>
        public override DocumentObject PreviousSibling
        {
            get
            {
                O.OpenXmlElement ele = XElement.PreviousSibling();
                if (ele == null) return null;
                return Construct(ele);
            }
        }

        /// <summary>
        /// Gets the object that immediately follows the current object. 
        /// Returns null if there is no next object.
        /// </summary>
        public override DocumentObject NextSibling
        {
            get
            {
                O.OpenXmlElement ele = XElement.NextSibling();
                if (ele == null) return null;
                return Construct(ele);
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Inserts the specified object immediately before the current object.
        /// </summary>
        /// <param name="obj">The new object to insert.</param>
        public override void InsertBeforeSelf(DocumentObject obj)
        {
            XElement.InsertBeforeSelf(obj.XElement);
        }

        /// <summary>
        /// Inserts the specified object immediately after the current object.
        /// </summary>
        /// <param name="obj">The new object to insert.</param>
        public override void InsertAfterSelf(DocumentObject obj)
        {
            XElement.InsertAfterSelf(obj.XElement);
        }

        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            O.OpenXmlElement ele = XElement.CloneNode(true);
            return Construct(ele);
        }

        /// <summary>
        /// Removes the current element from its parent.
        /// </summary>
        public override void Remove()
        {
            if (_element.Descendants<W.SectionProperties>().Any())
            {
                _element.RemoveAllChildren<W.Run>();
            }
            else
            {
                _element.Remove();
            }
        }
        #endregion


        #region Private Methods
        private IEnumerable<DocumentItem> ChildItems()
        {
            foreach(O.OpenXmlElement ele in _element.ChildElements)
            {
                if (ele is W.Paragraph)
                {
                    yield return new Paragraph(_doc, (W.Paragraph)ele);
                }
                else if (ele is W.Table)
                {
                    yield return new Table(_doc, (W.Table)ele);
                }
                else if(ele is W.SdtBlock)
                {
                    yield return new SdtBlock(_doc, (W.SdtBlock)ele);
                }
                else if (ele is W.Run)
                {
                    foreach (ParagraphItem item in RunItems((W.Run)ele))
                        yield return item;
                }
                else if (ele is W.Hyperlink)
                {
                    foreach (O.OpenXmlElement e in ele.ChildElements)
                    {
                        if (e is W.Run)
                        {
                            foreach (ParagraphItem item in RunItems((W.Run)e))
                                yield return item;
                        }
                    }
                }
                else if (ele is M.OfficeMath) // Office Math
                {
                    yield return new OfficeMath(_doc, (M.OfficeMath)ele);
                }
                else if (ele is M.Paragraph)
                {
                    foreach (M.OfficeMath oMath in ele.Elements<M.OfficeMath>())
                        yield return new OfficeMath(_doc, oMath);
                }
                else if(ele is W.DeletedRun) // DeletedRun
                {
                    yield return new DeletedRange(_doc, (W.DeletedRun)ele);
                }
                else if (ele is W.InsertedRun) // InsertedRun
                {
                    yield return new InsertedRange(_doc, (W.InsertedRun)ele);
                }
                else if(ele is W.SimpleField) // SimpleField
                {
                    yield return new SimpleField(_doc, (W.SimpleField)ele);
                }
                else if(ele is W.BookmarkStart) // BookmarkStart
                {
                    yield return new BookmarkStart(_doc, (W.BookmarkStart)ele);
                }
                else if (ele is W.BookmarkEnd) // BookmarkEnd
                {
                    yield return new BookmarkEnd(_doc, (W.BookmarkEnd)ele);
                }
            }
        }

        private IEnumerable<ParagraphItem> RunItems(W.Run run)
        {
            // text range
            if (run.Elements<W.Text>().Any())
                yield return new TextRange(_doc, run, run.GetFirstChild<W.Text>());

            // footnote reference
            if (run.Elements<W.FootnoteReference>().Any())
            {
                yield return new FootnoteReference(_doc, run, run.GetFirstChild<W.FootnoteReference>());
            }
            // endnote reference
            if (run.Elements<W.EndnoteReference>().Any())
            {
                yield return new EndnoteReference(_doc, run, run.GetFirstChild<W.EndnoteReference>());
            }
            // break
            if (run.Elements<W.Break>().Any())
            {
                yield return new Break(_doc, run, run.GetFirstChild<W.Break>());
            }
            // tab
            if (run.Elements<W.TabChar>().Any())
            {
                yield return new Tab(_doc, run, run.GetFirstChild<W.TabChar>());
            }
            // non breaking hyphen
            if (run.Elements<W.NoBreakHyphen>().Any())
            {
                yield return new NoBreakHyphen(_doc, run, run.GetFirstChild<W.NoBreakHyphen>());
            }
            // field char
            if (run.Elements<W.FieldChar>().Any())
            {
                yield return new FieldChar(_doc, run, run.GetFirstChild<W.FieldChar>());
            }
            // field code
            if (run.Elements<W.FieldCode>().Any())
            {
                yield return new FieldCode(_doc, run, run.GetFirstChild<W.FieldCode>());
            }
            // drawing
            foreach (W.Drawing drawing in run.Descendants<W.Drawing>())
            {
                A.GraphicData graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
                if (graphicData != null)
                {
                    if (graphicData.FirstChild is Pic.Picture)
                        yield return new Picture(_doc, run, drawing);
                    else if (graphicData.FirstChild is Wps.WordprocessingShape)
                        yield return new Shape(_doc, run, drawing);
                    else if (graphicData.FirstChild is Wpg.WordprocessingGroup)
                        yield return new GroupShape(_doc, run, drawing);
                    else if (graphicData.FirstChild is Wpc.WordprocessingCanvas)
                        yield return new Canvas(_doc, run, drawing);
                    else if (graphicData.FirstChild is Dgm.RelationshipIds)
                        yield return new Diagram(_doc, run, drawing);
                    else if (graphicData.FirstChild is C.ChartReference)
                        yield return new Chart(_doc, run, drawing);
                }
            }
            // vml picture
            if (run.Elements<W.Picture>().Any())
            {
                yield return new Picture(_doc, run, run.GetFirstChild<W.Picture>());
            }
            // embedded object
            foreach (W.EmbeddedObject obj in run.Elements<W.EmbeddedObject>())
            {
                yield return new EmbeddedObject(_doc, run, obj);
            }
            // deleted text range
            if (run.Elements<W.DeletedText>().Any())
            {
                yield return new DeletedTextRange(_doc, run);
            }
        }

        private DocumentObject Construct(O.OpenXmlElement element)
        {
            if (element is W.Paragraph) return new Paragraph(_doc, (W.Paragraph)element);
            else if (element is W.Table) return new Table(_doc, (W.Table)element);
            else if (element is W.SdtBlock) return new SdtBlock(_doc, (W.SdtBlock)element);
            else if (element is W.SdtProperties) return new SdtProperties(_doc, (W.SdtProperties)element);
            else if (element is W.SdtContentBlock) return new SdtContent(_doc, (W.SdtContentBlock)element);
            else return null;
        }
        #endregion
    }
}
