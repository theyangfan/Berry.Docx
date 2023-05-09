using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class SdtBlockGenerator
    {
        public static SdtBlock Generate(Document doc)
        {
            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();
            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtContentBlock1);
            return sdtBlock1;
        }
    }
}
