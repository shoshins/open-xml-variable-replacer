using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.VariableReplacer
{
    public interface IVariableReplacer
    {
        void ReplaceVariables(Document document, Dictionary<string, string> replacers,
            string openVariableSymbol = "{$", string closeVariableSymbol = "$}");
    }
}