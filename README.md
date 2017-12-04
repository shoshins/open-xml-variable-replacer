# open-xml-variable-replacer
Simple and lightweight replacer for varibles in OpenXml (DOCX MS Word).

# Installation (via Nuget)
```
Install-Package DocumentFormat.OpenXml.VariableReplacer -Version 0.1.0
```
# Usage

1) Register or initialize:
```
IVariableReplacer _openXmlVariableReplacer = new VariableReplacer();
```
or some DI like this:
```
serviceCollection.AddTransient<IVariableReplacer, VariableReplacer>();
```

2) Use:
```
using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
{
    MainDocumentPart mainPart = document.MainDocumentPart;
    Dictionary<string, string> replacers = new Dictionary<string, string> { 
      {"MyVariable", "MyNewText"},
      {"MyVariable2", "MyNewText2"},
      {"MyVariable3", "MyNewText3"},
      // ...
    };
    _openXmlVariableReplacer.ReplaceVariables(mainPart.Document, replacers);
}
```
