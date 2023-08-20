# DocXReducer

**DocXReducer** is a library for reducing size of Word documents **(.docx)** by simplifying theirs XML. Consider using it for archived documents, because after reducing some Word functions may not work correctly

This library uses only `DocumentFormat.OpenXml`

## Get Started

Install package `DocxReducer` with the Nuget package manager or the CLI

```
dotnet add package DocxReducer
```

Then with the class
```cs
var reducer = new Reducer();
```

it's possible to process a document by path

```cs
reducer.Reduce(@".\Document.docx");
```

or with a WordprocessingDocument class
```cs
var docx = WordprocessingDocument.Open(@".\Document.docx", true);

reducer.Reduce(docx);
```

## Build console app

To build console app enter:

```bash
dotnet publish -c Release -r rid
```

where `rid` is Runtime Identifier of target system. Possible values are:

- Windows: `win-x64`
- Linux: `linux-x64`
- MacOS: `osx-x64`

Full list of rids can be found [here](https://github.com/dotnet/runtime/blob/main/src/libraries/Microsoft.NETCore.Platforms/src/runtime.json)
