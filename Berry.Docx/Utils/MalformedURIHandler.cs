using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace Berry.Docx.Utils
{
    /// <summary>
    /// Handle malformed hyperlink.
    /// </summary>
    internal class MalformedURIHandler : RelationshipErrorHandler
    {
        public override string Rewrite(Uri partUri, string id, string uri)
        {
            // return the correct hyperlink
            return Regex.Match(uri, @"https?://(\w+\.?)+")?.Value;
        }
    }
}
