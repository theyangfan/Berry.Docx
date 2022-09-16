using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class InsertedRange : ParagraphItem
    {
        #region Priavate Members
        private readonly W.InsertedRun _ins;
        #endregion

        #region Constructors
        internal InsertedRange(Document doc, W.InsertedRun ins) : base(doc, ins)
        {
            _ins = ins;
        }

        #endregion

        #region Public Properties
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.InsertedRange;

        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                foreach (var item in ChildObjects)
                {
                    if (item is TextRange)
                    {
                        text.Append(((TextRange)item).Text);
                    }
                }
                return text.ToString();
            }
            set
            {
                _ins.RemoveAllChildren<W.Run>();
                W.Run run = RunGenerator.Generate(value);
                _ins.Append(run);
            }
        }
        #endregion
    }
}
