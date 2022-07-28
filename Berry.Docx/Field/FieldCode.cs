using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;

namespace Berry.Docx.Field
{
    internal class FieldCode
    {
        private List<O.OpenXmlElement> _childElements = new List<O.OpenXmlElement>();
        private string _code = "";
        private string _result = "";
        public FieldCode(List<O.OpenXmlElement> childElements)
        {
            _childElements = childElements;
            bool begin = false;
            bool separate = false;
            int begin_times = 0;
            int invalid_begin_times = 0;
            foreach (O.OpenXmlElement ele in childElements)
            {
                if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.SimpleField"))
                {
                    W.SimpleField fldSimple = (W.SimpleField)ele;
                    _code += "{" + fldSimple.Instruction.Value + "}";
                    _result += fldSimple.InnerText;
                }
                else if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.Run"))
                {
                    W.Run run = (W.Run)ele;
                    if (run.Ancestors<W.SimpleField>().Any())
                        continue;
                    if (run.Elements<W.FieldChar>().Any() && run.Elements<W.FieldChar>().First().FieldCharType != null)
                    {
                        string field_type = run.Elements<W.FieldChar>().First().FieldCharType.ToString();
                        if (field_type == "begin")
                        {
                            if (!separate)
                            {
                                begin_times++;
                                _code += "{";
                                begin = true;
                                continue;
                            }
                            else
                            {
                                ++invalid_begin_times;
                            }

                        }
                        if (field_type == "separate")
                        {
                            separate = true;
                            continue;
                        }
                        if (field_type == "end")
                        {
                            if (invalid_begin_times == 0)
                            {
                                _code += "}";
                                if (--begin_times == 0)
                                {
                                    begin = false;
                                }
                                separate = false;
                                continue;
                            }
                            else
                            {
                                --invalid_begin_times;
                            }
                        }
                    }
                    if (begin && !separate)
                    {
                        _code += run.InnerText;
                    }
                    if (begin && separate)
                    {
                        _result += run.InnerText;
                    }
                }
            }
        }

        public FieldCode(W.SimpleField simpleFld)
        {
            _code = "{" + simpleFld.Instruction.Value + "}";
            _result = simpleFld.InnerText;
        }

        public string Code { get => _code; }
        public string Result { get => _result; }

        public FieldCodeCollection ChildFieldCodes
        {
            get
            {
                List<FieldCode> fieldcodes = new List<FieldCode>();
                List<O.OpenXmlElement> childElements = new List<O.OpenXmlElement>();

                int begin_times = 0;
                int end_times = 0;

                foreach (O.OpenXmlElement ele in _childElements)
                {
                    if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.SimpleField"))
                    {
                        fieldcodes.Add(new FieldCode((W.SimpleField)ele));
                    }
                    else if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.Run"))
                    {
                        W.Run run = (W.Run)ele;
                        if (run.Elements<W.FieldChar>().Any() && run.Elements<W.FieldChar>().First().FieldCharType != null)
                        {
                            string field_type = run.Elements<W.FieldChar>().First().FieldCharType.ToString();
                            if (field_type == "begin" && ++begin_times == 1) continue;
                            else if (field_type == "end" && ++end_times == begin_times) continue;
                        }
                        if (begin_times > 0)
                        {
                            childElements.Add(ele);
                            if (end_times == begin_times)
                            {
                                fieldcodes.Add(new FieldCode(childElements));
                                begin_times = 0;
                                end_times = 0;
                                childElements.Clear();
                            }
                        }
                    }
                }
                return new FieldCodeCollection(fieldcodes);
            }
        }
    }
}
