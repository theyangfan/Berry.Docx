using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Berry.Docx.Documents;

namespace Berry.Docx.Interface
{
    public interface IDocumentObject
    {
        //TODO
        DocumentObject Owner { get; }
        IDocumentObject PreviousSibling { get; }
        IDocumentObject NextSibling { get; }
    }
}
