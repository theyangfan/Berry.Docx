﻿using System;
using System.Collections.Generic;
using System.Linq;
using O = DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a paragraph child item.
    /// </summary>
    public abstract class ParagraphItem : DocumentItem
    {
        #region Private Members
        private readonly Document _doc;
        // the owner openxml run element.
        private readonly W.Run _ownerRun;
        private readonly O.OpenXmlElement _element;
        #endregion

        #region Constructors
        /// <summary>
        /// When the ele is a part of run element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerRun"></param>
        /// <param name="ele"></param>
        internal ParagraphItem(Document doc,  W.Run ownerRun, O.OpenXmlElement ele)
            : base(doc, ele)
        {
            _doc = doc;
            _ownerRun = ownerRun;
        }
        /// <summary>
        /// When the ele is not a part of run element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ele"></param>
        internal ParagraphItem(Document doc, O.OpenXmlElement ele)
            : base(doc, ele)
        {
            _doc = doc;
            _element = ele;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the parent paragraph of the current item.
        /// </summary>
        public Paragraph OwnerParagraph
        {
            get
            {
                if(_ownerRun != null)
                    return new Paragraph(_doc, _ownerRun.Ancestors<W.Paragraph>().First());
                else
                    return new Paragraph(_doc, _element.Ancestors<W.Paragraph>().First());
            }
        }
        
        #endregion

        #region Public Methods
        /// <summary>
        /// Appends a comment to the current paragraph item.
        /// </summary>
        /// <param name="author">The author of the comment.</param>
        /// <param name="contents">The paragraphs content of the comment.</param>
        public void AppendComment(string author, params string[] contents)
        {
            int id = 0; // comment id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // max id + 1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // comments content

            W.Comment comment = new W.Comment() { Id = id.ToString(), Author = author };
            foreach (string content in contents)
            {
                W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
                comment.Append(paragraph);
            }
            comments.Append(comment);
            // comment mark
            W.CommentRangeStart startMark = new W.CommentRangeStart() { Id = id.ToString() };
            W.CommentRangeEnd endMark = new W.CommentRangeEnd() { Id = id.ToString() };
            W.Run referenceRun = new W.Run(new W.CommentReference() { Id = id.ToString() });
            // Insert comment mark
            if(_ownerRun != null)
            {
                _ownerRun.InsertBeforeSelf(startMark);
                _ownerRun.InsertAfterSelf(endMark);
            }
            else
            {
                _element.InsertBeforeSelf(startMark);
                _element.InsertAfterSelf(endMark);
            }
            endMark.InsertAfterSelf(referenceRun);
        }

        /// <summary>
        /// Appends a comment to the current paragraph item.
        /// </summary>
        /// <param name="author">The author of the comment.</param>
        /// <param name="contents">The paragraphs content of the comment.</param>
        public void AppendComment(string author, IEnumerable<string> contents)
        {
            int id = 0; // comment id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // max id + 1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // comments content
            W.Comment comment = new W.Comment() { Id = id.ToString(), Author = author };
            foreach (string content in contents)
            {
                W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
                comment.Append(paragraph);
            }
            comments.Append(comment);
            // comment mark
            W.CommentRangeStart startMark = new W.CommentRangeStart() { Id = id.ToString() };
            W.CommentRangeEnd endMark = new W.CommentRangeEnd() { Id = id.ToString() };
            W.Run referenceRun = new W.Run(new W.CommentReference() { Id = id.ToString() });
            // Insert comment mark
            if (_ownerRun != null)
            {
                _ownerRun.InsertBeforeSelf(startMark);
                _ownerRun.InsertAfterSelf(endMark);
            }
            else
            {
                _element.InsertBeforeSelf(startMark);
                _element.InsertAfterSelf(endMark);
            }
            endMark.InsertAfterSelf(referenceRun);
        }
        #endregion
    }
}