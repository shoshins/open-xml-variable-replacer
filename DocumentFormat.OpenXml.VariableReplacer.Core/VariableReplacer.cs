using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.VariableReplacer.Core
{
    public class VariableReplacer: IVariableReplacer
    {
        public void ReplaceVariables(Document document, Dictionary<string, string> replacers,
            string openVariableSymbol = "{$", string closeVariableSymbol = "$}")
        {
            // Get Document lines
            IEnumerable<Paragraph> paragraphs = document.Body.Descendants<Paragraph>();
            // Collections with variables that we will found.
            IDictionary<Guid, ICollection<Text>> variableItems = new Dictionary<Guid, ICollection<Text>>();
            // Determines whether Variable opener was found. We will save all text fields while close symbol not found.
            bool isVariable = false;
            // All found text fields list will be part of variable that have personal identifier.
            Guid variableId = Guid.NewGuid();

            // Go around all documents lines
            foreach (Paragraph para in paragraphs)
            {
                // Take all Run fragments from line
                foreach (Run run in para.Elements<Run>())
                {
                    // Take all text fragments from line
                    foreach (Text text in run.Elements<Text>())
                    {
                        // Check if this fragment contains variable start. Also, we check that at this moment no one another variable is started.
                        if (text.Text.Contains(openVariableSymbol) && !isVariable)
                        {
                            // If this is fragment with open symbol, we will save it. And start process of searching close symbol.
                            variableId = Guid.NewGuid();
                            variableItems.Add(variableId, new List<Text> {text});
                            isVariable = true;
                        }
                        else
                        {
                            if (text.Text.Contains(closeVariableSymbol) && isVariable)
                            {
                                // If this is fragment with close symbol, we will save it. And stop process of searching variable parts.
                                if (variableItems.TryGetValue(variableId, out ICollection<Text> items))
                                {
                                    items.Add(text);
                                }
                                isVariable = false;
                            }
                            else if (isVariable)
                            {
                                // This is not a start or close fragment. It is just a part of variable, save it.
                                if (variableItems.TryGetValue(variableId, out ICollection<Text> items))
                                {
                                    items.Add(text);
                                }
                            }
                        }
                    }
                }
            }

            // Here we found all what we need. At this point we will 
            // 1) Foreach all variables objects and join it to single line
            // 2) Get replacer by this line and replace variable.
            // 3) In the opener fragment we will have new replaced text right now.
            // 4) Remove all variable fragments except opener one. Variable fully replaced.

            foreach (KeyValuePair<Guid, ICollection<Text>> variableDescription in variableItems)
            {
                StringBuilder variableBuilder = new StringBuilder();
                Text variableOpenerText = null;

                // Joining variable items in single line. And removing all except the first one with opener symbol.
                foreach (Text variable in variableDescription.Value)
                {
                    variableBuilder.Append(variable.Text);
                    if (variable.Text.Contains(openVariableSymbol))
                    {
                        if (variableOpenerText != null)
                        {
                            variable.Remove();
                        }
                        else
                        {
                            variableOpenerText = variable;
                        }
                    }
                    else
                    {
                        variable.Remove();
                    }
                }
                string buildedVariable = variableBuilder.ToString();

                // Replace variable with replacer by input dictionary.

                Regex regex = new Regex("{\\$(.*)\\$}");
                Match match = regex.Match(buildedVariable);
                if (match.Success && match.Groups.Count > 1)
                {
                    string variableText = match.Groups[1].Value;
                    if (replacers.TryGetValue(variableText, out string variableReplacer))
                    {
                        if (variableOpenerText != null)
                        {
                            variableOpenerText.Text =
                                variableOpenerText.Text.Replace(openVariableSymbol, variableReplacer);
                        }
                    }
                    else
                    {
                        variableOpenerText?.Remove();
                    }
                }
            }
        }
    }
}