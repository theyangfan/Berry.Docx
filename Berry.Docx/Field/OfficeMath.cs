using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using M = DocumentFormat.OpenXml.Math;

namespace Berry.Docx.Field
{
    public class OfficeMath : ParagraphItem
    {
        private readonly M.OfficeMath _oMath;
        private readonly M.Paragraph _oMathPara;
        internal OfficeMath(Document doc, M.OfficeMath oMath) : base(doc, oMath)
        {
            _oMath = oMath;
            _oMathPara = oMath.Ancestors<M.Paragraph>().FirstOrDefault();
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.OfficeMath;

        public bool IsInline()
        {
            return _oMathPara == null;
        }

        public OfficeMathJustificationType Justification
        {
            get
            {
                if(_oMathPara != null)
                {
                    M.Justification jc = _oMathPara.ParagraphProperties?.Justification;
                    if (jc != null)
                    {
                        if (jc.Val.Value == M.JustificationValues.Left)
                            return OfficeMathJustificationType.Left;
                        else if (jc.Val.Value == M.JustificationValues.Right)
                            return OfficeMathJustificationType.Right;
                        else if (jc.Val.Value == M.JustificationValues.Center)
                            return OfficeMathJustificationType.Center;
                        else
                            return OfficeMathJustificationType.CenterGroup;
                    }
                    return OfficeMathJustificationType.CenterGroup;
                }
                return OfficeMathJustificationType.Invalid;
            }
        }

    }
}
