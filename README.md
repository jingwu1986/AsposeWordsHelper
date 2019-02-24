# AsposeWordsHelper
A word writer helper/wrapper for "aspose.words", including title, paragraph, break, table, image and ole object. It simplifies the code to generate a word file. For example, to merge table cells, it uses WordTableMergeOption to do that, instead of writing a few lines of code to implement it.

# Usage
```
WordDocumentOption documentOption = new WordDocumentOption() { NeedTableOfContents = false, AutoNumberTitle = true, AutoNumberType = AutoNumberType.EN, TableFontSize = 10, InsertParagraphBreakWhenChildrenEnd = true };
List<WordNode> nodes = new List<WordNode>();

WordTitleOption level1TitleOption = new WordTitleOption() { StyleIdentifier = StyleIdentifier.Heading1 };
WordTitleOption level2TitleOption = new WordTitleOption() { StyleIdentifier = StyleIdentifier.Heading2 };
WordTitleOption level3TitleOption = new WordTitleOption() { StyleIdentifier = StyleIdentifier.Heading3 };

nodes.Add(
    new WordParagraphNode("Sample Document", new WorldParagraphOption { Alignment = ParagraphAlignment.Center, FontSize = 15, Bold = true })
    );

nodes.Add(
    new WordTitleNode("Title 1", level1TitleOption,
    new WordTitleNode("Title 1.1", level2TitleOption)
));

nodes.Add(
    new WordTitleNode("Table 1", level1TitleOption,
    new WordTableNode(GetWordTable1())
));

nodes.Add(
    new WordTitleNode("Table 2", level1TitleOption,
    new WordTableNode(GetWordTable2())
));

nodes.Add(
    new WordTitleNode("Table of auto merge", level1TitleOption,
    new WordTableNode(GetAutoMergeTable())
));

nodes.Add(
    new WordTitleNode("Table of manually merge", level1TitleOption,
    new WordTableNode(GetManuallyMergeTable())
));

nodes.Add(
        new WordTitleNode("Image", level1TitleOption,
        new WordImageNode(GetWordImage())
));

nodes.Add(
    new WordTitleNode("Ole Object", level1TitleOption,
    new WordOleObjectNode(GetWordOleObject())
));

string filePath = "test.docx";
WordWriter writer = new WordWriter(nodes, documentOption);

writer.Write(filePath);
```

# Screenshort of generated file
![Screenshort](https://github.com/jingwu1986/AsposeWordsHelper/blob/master/Screenshort.png)
