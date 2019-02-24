using Aspose.Words;
using AsposeWordsHelper.Sample.Properties;
using System.Collections.Generic;
using System.Diagnostics;

namespace AsposeWordsHelper.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
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

            Process.Start(filePath);
        }       

        private static WordTable GetWordTable1()
        {
            WordTable table = new WordTable() { ShowHeader = true };
            table.Columns = new List<WordTableColumn>()
            {
                new WordTableColumn("Name") { Width = 20, IsRowHeader = false, Alignment = ParagraphAlignment.Center },              
                new WordTableColumn("Description") { Width = 80, IsRowHeader = false, Alignment = ParagraphAlignment.Left }                
            };
            table.Data = new List<dynamic>()
            {
                new
                {
                    Name = "CPU",
                    Description = "Central Processing Unit"
                },
                new
                {
                    Name = "SSD",
                    Description = "Solid State Disk"
                }
            };           
            return table;
        }

        private static WordTable GetWordTable2()
        {
            WordTable table = new WordTable() { ShowHeader = false };
            table.Columns = new List<WordTableColumn>()
            {
                new WordTableColumn("Name1") { Width = 20, IsRowHeader = true, Alignment = ParagraphAlignment.Center },
                new WordTableColumn("Description1") { Width = 50, Alignment = ParagraphAlignment.Left },
                new WordTableColumn("Name2") { Width = 20, IsRowHeader = true, Alignment = ParagraphAlignment.Center },
                new WordTableColumn("Description2") { Width = 50, Alignment = ParagraphAlignment.Left }
            };
            table.Data = new List<dynamic>()
            {
                new
                {
                    Name1 = "CPU",
                    Description1 = "Central Processing Unit",
                    Name2 = "SSD",
                    Description2 = "Solid State Disk",
                },
                new
                {
                    Name1 = "USB",
                    Description1 = "Universal Serial Bus",
                    Name2 = "GPU",
                    Description2 = "Graphics Processing Unit",
                }
            };
           
            return table;
        }

        private static WordTable GetAutoMergeTable()
        {
            WordTable table = new WordTable() { ShowHeader = true };
            table.Columns = new List<WordTableColumn>()
            {
                new WordTableColumn("Category") { Width = 20,  Alignment = ParagraphAlignment.Center, AutoMerge=true },
                new WordTableColumn("Name") { Width = 20, Alignment = ParagraphAlignment.Center },
                new WordTableColumn("Description") { Width = 60, Alignment = ParagraphAlignment.Left }
            };
            table.Data = new List<dynamic>()
            {
                new
                {
                    Category="Computer",
                    Name = "CPU",
                    Description = "Central Processing Unit"
                },
                new
                {
                    Category="Computer",
                    Name = "GPU",
                    Description = "Solid State Disk"
                }
            };
            return table;
        }

        private static WordTable GetManuallyMergeTable()
        {
            WordTable table = new WordTable() { ShowHeader = true };
            table.Columns = new List<WordTableColumn>()
            {
                new WordTableColumn("Category") { Width = 20,  Alignment = ParagraphAlignment.Center },
                new WordTableColumn("Name") { Width = 20, Alignment = ParagraphAlignment.Center },
                new WordTableColumn("Description") { Width = 60, Alignment = ParagraphAlignment.Left }
            };
            table.Data = new List<dynamic>()
            {
                new
                {
                    Category="Computer",
                    Name = "CPU",
                    Description = "Central Processing Unit"
                },
                new
                {
                    Category="Computer",
                    Name = "GPU",
                    Description = "Solid State Disk"
                },
                new
                {
                    Category="Computer",
                    Name = "USB",
                    Description = "USB"
                }
            };

            table.MergeOptions = new List<WordTableMergeOption>() {
                new WordTableMergeOption() {  StartColumnIndex=0, StartRowIndex=0, Rowspan=2},
                new WordTableMergeOption() {  StartColumnIndex=1, StartRowIndex=2, Colspan=2}
            };

            return table;
        }

        private static WordImage GetWordImage()
        {
            return new WordImage(Resources.test) { };
        }

        private static WordOleObject GetWordOleObject()
        {           
            string filePath = "sample.xlsx";

            return new WordOleObject("sample file", filePath, true, Resources.excel) { DeleteFileAfterUsed = false };
        }
    }
}
