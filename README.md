# DocXReducer

**DocXReducer** is a library for reducing size of Word documents **(.docx)** by simplifying theirs XML

This library uses only `DocumentFormat.OpenXml 2.14.0`

## Usage

There are two ways of using reducer:

- with a file name
```cs
var reducer = new Reducer();
reducer.Reduce(@".\Document.docx");
```

- or with a WordprocessingDocument class
```cs
var inputFilePath = @".\Document.docx";

var docx = WordprocessingDocument.Open(inputFilePath, true);

var reducer = new Reducer();

reducer.Reduce(docx);
```

## Constructor options

`Reducer` class has some options
```cs
public Reducer(bool deleteBookmarks=true,
               bool createNewStyles=true)
```

| Name | Default value | Description | Notes |
|-----------------| ----- | ------------------ | ------------------ |
| deleteBookmarks | true | Delete or not bookmarks in the document | ... |
| createNewStyles | true | Replace run properties with styles or not | In some cases file can be bigger with `true` |

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
