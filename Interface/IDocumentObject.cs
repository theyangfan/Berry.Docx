using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx.Interface
{
    public interface IDocumentObject
    {
        DocumentObject Owner { get; }
        IDocumentObject PreviousSibling { get; }
        IDocumentObject NextSibling { get; }
    }
}
