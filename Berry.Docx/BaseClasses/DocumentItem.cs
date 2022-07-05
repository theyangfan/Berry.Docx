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
        public override DocumentObjectCollection ChildObjects
        {
            get
            {
                return new DocumentItemCollection(_element, Children());
            }
        }
        #endregion

        #region Public Methods
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

        public override DocumentObject Clone()
        {
            O.OpenXmlElement ele = XElement.CloneNode(true);
            if (ele is W.Paragraph) return new Paragraph(_doc, (W.Paragraph)ele);
            else if (ele is W.Table) return new Table(_doc, (W.Table)ele);
            else if (ele is W.SdtBlock) return new SdtBlock(_doc, (W.SdtBlock)ele);

            return null;
        }
        #endregion


        #region Private Methods
        private IEnumerable<DocumentItem> Children()
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
                    yield return new OfficeMath(_doc, ele as M.OfficeMath);
                }
                else if (ele is M.Paragraph)
                {
                    foreach (M.OfficeMath oMath in ele.Elements<M.OfficeMath>())
                        yield return new OfficeMath(_doc, oMath);
                }
            }
        }

        private IEnumerable<ParagraphItem> RunItems(W.Run run)
        {
            // text range
            if (run.Elements<W.Text>().Any())
                yield return new TextRange(_doc, run);

            // footnote reference
            if (run.Elements<W.FootnoteReference>().Any())
            {
                yield return new FootnoteReference(_doc, run, run.Elements<W.FootnoteReference>().First());
            }
            // endnote reference
            if (run.Elements<W.EndnoteReference>().Any())
            {
                yield return new EndnoteReference(_doc, run, run.Elements<W.EndnoteReference>().First());
            }
            // break
            if (run.Elements<W.Break>().Any())
            {
                yield return new Break(_doc, run, run.Elements<W.Break>().First());
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
                yield return new Picture(_doc, run, run.Elements<W.Picture>().First());
            }
            // embedded object
            foreach (W.EmbeddedObject obj in run.Elements<W.EmbeddedObject>())
            {
                yield return new EmbeddedObject(_doc, run, obj);
            }
        }
        #endregion
    }
}
