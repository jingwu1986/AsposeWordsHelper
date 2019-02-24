using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;
using Aspose.Words.Drawing;
using System.Drawing.Drawing2D;

namespace AsposeWordsHelper
{
    public class WordGenerator
    {
        #region Fields & Properties
        private string templateFile;
        private WordDocumentOption option;
        private Document document { get; set; }
        private DocumentBuilder builder { get; set; }
        private bool hasAddedTableOfContents = false;

        public Document Document
        {
            get
            {
                return this.document;
            }
        }

        public DocumentBuilder Builder
        {
            get
            {
                return this.builder;
            }
        }

        public WordDocumentOption Option
        {
            get
            {
                return this.option;
            }
        }
        #endregion

        public WordGenerator()
        {
            this.Init();
            this.Prepare();
        }

        public WordGenerator(WordDocumentOption option)
        {
            this.Init();
            this.option = option;
            this.Prepare();
        }

        public WordGenerator(string templateFile, WordDocumentOption option)
        {
            this.templateFile = templateFile;
            this.Init();
            this.option = option;
            this.Prepare();
        }

        private void Init()
        {
            if(!string.IsNullOrEmpty(this.templateFile) && File.Exists(this.templateFile))
            {
                this.document = new Document(this.templateFile);
            }
            else
            {
                this.document = new Document();
            }
           
            this.builder = new DocumentBuilder(this.document);
            this.document.FirstSection.Body.PrependChild(new Paragraph(this.document));
            this.builder.MoveToDocumentStart();

            this.InitDefaultOptions();
            this.InitDefaultFormat();
        }

        private void Prepare()
        {
            if (this.option.NeedTableOfContents)
            {
                this.AddTableOfContents();
            }
        }

        private void InitDefaultOptions()
        {
            this.option = new WordDocumentOption();

            PageSetup ps = this.builder.PageSetup;
            ps.LeftMargin = ConvertUtil.InchToPoint(0.4);
            ps.RightMargin = ConvertUtil.InchToPoint(0.4);
            ps.PaperSize = PaperSize.A4;
            ps.Orientation = Orientation.Portrait;

            this.option.PageSetup = ps;
        }

        private void InitDefaultFormat()
        {
            this.InitFont();
            this.builder.RowFormat.HeightRule = HeightRule.Exactly;
            this.builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            this.builder.ParagraphFormat.LineSpacing = this.option.LineSpacing;
        }

        private void InitFont()
        {
            this.builder.Font.ClearFormatting();
            this.builder.Font.Name = this.option.FontName;
            this.builder.Font.Size = this.option.FontSize;           
        }

        private void Complete()
        {
            this.document.UpdateFields();
        }
      
        public void SaveTo(Stream stream)
        {
            this.Complete();
            this.document.Save(stream, this.option.SaveFormat);
            stream.Position = 0;
        }
      
        public void SaveTo(string filePath)
        {
            this.Complete();
            this.document.Save(filePath, this.option.SaveFormat);
        }
      
        public void AddTableOfContents()
        {
            if (!this.hasAddedTableOfContents)
            {
                this.builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
                this.hasAddedTableOfContents = true;
            }
        }
      
        public void AddTitle(string title, WordTitleOption option = null)
        {
            this.ClearParagraphFormat();            

            if (option != null)
            {
                this.builder.ParagraphFormat.StyleIdentifier = option.StyleIdentifier;
                this.builder.Font.Size = WordUtil.PickValue(option.FontSize, this.builder.Font.Size);
                this.builder.Font.Bold = option.Bold;
                this.builder.Font.Italic = option.Italic;
                this.builder.ParagraphFormat.Alignment = option.Alignment;
                if (option.LeftIndent > 0)
                {
                    this.builder.ParagraphFormat.LeftIndent = option.LeftIndent;
                }
            }

            this.builder.Writeln(title ?? "");

            this.InitFont();
        }

        public void AddPageBreak()
        {
            this.ClearParagraphFormat();
            this.builder.InsertBreak(BreakType.PageBreak);
        }

        public void AddLineBreak()
        {
            this.ClearParagraphFormat();
            this.builder.InsertBreak(BreakType.LineBreak);          
        }

        public void AddParagraphBreak()
        {
            this.ClearParagraphFormat();
            this.builder.InsertBreak(BreakType.ParagraphBreak);
        }

        private void ClearParagraphFormat()
        {
            this.builder.ParagraphFormat.ClearFormatting();                    
        }

        private void ClearFontFormat()
        {
            this.builder.Font.ClearFormatting();
        }
      
        public void AddParagraph(string content, WorldParagraphOption option = null)
        {
            double leftIndent = this.builder.ParagraphFormat.LeftIndent;
            this.ClearParagraphFormat();         
            
            if (option != null)
            {
                this.builder.Font.Size = option.FontSize;
                this.builder.ParagraphFormat.Alignment = option.Alignment;
                this.builder.Font.Bold = option.Bold;
                this.builder.Font.Italic = option.Italic;

                double actualLeftIndent = 0;
                if (option.LeftIndent > 0)
                {
                    actualLeftIndent = option.LeftIndent;                   
                }
                else
                {
                    if (this.option.ParagraphLeftIndentAsBefore)
                    {
                        actualLeftIndent = leftIndent;
                    }
                }

                if(actualLeftIndent>0)
                {
                    this.builder.ParagraphFormat.LeftIndent = actualLeftIndent;
                }

                double firstLineIndent = WordUtil.PickValue(option.FirstLineIndent, this.option.ParagraphFirstLineIndent);
                if (firstLineIndent > 0)
                {
                    this.builder.ParagraphFormat.FirstLineIndent = actualLeftIndent + firstLineIndent;
                }
            }

            if (option != null && option.IsHtml)
            {
                this.builder.InsertHtml(content ?? "", option.HtmlUseBuilderFormatting);
                this.builder.InsertBreak(BreakType.ParagraphBreak);
            }
            else
            {
                this.builder.Writeln(content ?? "");
            }
            
            this.InitFont();
        }

        #region Table      
        public void AddTable(WordTable table)
        {
            double leftIndent = this.builder.ParagraphFormat.LeftIndent;
            this.ClearParagraphFormat();           
            Table tb = this.builder.StartTable();

            int columnCount = table.Columns.Count;
            bool hasCreatedRow = false;

            if (table.TableWidthIsPercent)
            {
                table.ActualWidth = this.option.PageWidth * ((table.Width==0? this.option.TableDefaultWidth: table.Width) /100);                
            }
            else
            {
                table.ActualWidth = table.Width == 0 ? this.option.PageWidth : table.Width;
            }            

            double rowHeight = table.AutoRowHeight ? this.builder.RowFormat.Height : (table.RowHeight == 0 ? this.option.TableRowHeight : table.RowHeight);
            this.builder.RowFormat.Height = rowHeight;

            #region Header
            foreach (WordTableColumn column in table.Columns)
            {
                column.ActualWidth = WordUtil.CalculateTableColumnWidth(table, column);
                column.FontSize = WordUtil.PickValue(column.FontSize, this.option.TableFontSize, this.option.FontSize);
                if (table.ShowHeader)
                {
                    hasCreatedRow = true;
                    this.CreateTableHeaderCell(column);
                }
            }

            if (hasCreatedRow)
            {
                this.builder.EndRow();
            }
            #endregion

            #region Data
            if (table.Data != null)
            {
                if (!(table.Data is IEnumerable<object>))
                {
                    table.Data = new List<dynamic>() { table.Data };
                }
              
                int rowIndex =0;
                Dictionary<int, Dictionary<string, object>> dictRows = new Dictionary<int, Dictionary<string, object>>();
                foreach (dynamic row in table.Data)
                {
                    int rownumber = rowIndex + 1;

                    Dictionary<string, object> dict = WordUtil.ConvertToDictionary(row);
                    dictRows.Add(rownumber, dict);

                    int columnIndex = 0;                    

                    WordTableMergeOption mergeOption = null;            

                    foreach (WordTableColumn column in table.Columns)
                    {
                        object value = column.IsRowNumber ?
                          rownumber.ToString() :
                          (dict.ContainsKey(column.Name) ? dict[column.Name] : string.Empty);

                        #region Merge Cells

                        mergeOption = WordUtil.GetTableMergeOption(table, rowIndex, columnIndex);
                        if (mergeOption != null)
                        {
                            int colspan = mergeOption.Colspan;
                            int rowspan = mergeOption.Rowspan;

                            bool isHorizontalMergeStart = mergeOption.StartColumnIndex == columnIndex;
                            bool isVerticalMergeStart = mergeOption.StartRowIndex == rowIndex;
                            bool isLastMerge = (columnIndex == mergeOption.StartColumnIndex + colspan - 1) && (rowIndex == mergeOption.StartRowIndex + rowspan - 1);

                            this.AddTableCell(column, value, rowIndex, dictRows);

                            this.builder.CellFormat.HorizontalMerge = isHorizontalMergeStart? CellMerge.First:CellMerge.Previous;
                            this.builder.CellFormat.VerticalMerge = isVerticalMergeStart ? CellMerge.First : CellMerge.Previous;

                            if(isLastMerge)
                            {
                                mergeOption = null;
                            }

                            columnIndex++;
                            continue;
                        }
                        #endregion
                        else
                        {
                            this.AddTableCell(column, value, rowIndex, dictRows);
                        }                 

                        columnIndex++;
                    }
                    this.builder.EndRow();
                    rowIndex++;
                }
            }
            #endregion

            #region Footer
            if (table.Footer != null)
            {
                this.builder.RowFormat.Height = rowHeight;
                List<Cell> footerCells = new List<Cell>();

                LineStyle oldLeftLineStyle = this.builder.RowFormat.Borders.Left.LineStyle;
                LineStyle oldRightLineStyle = this.builder.RowFormat.Borders.Left.LineStyle;
                LineStyle oldBottomLineStyle = this.builder.RowFormat.Borders.Left.LineStyle;

                footerCells.Add(this.builder.InsertCell());

                if (table.Footer.IsHtml)
                {
                    this.builder.InsertHtml(table.Footer.Content);
                }
                else
                {
                    this.builder.Write(table.Footer.Content);
                }

                this.builder.CellFormat.HorizontalMerge = CellMerge.First;
                for (int i = 1; i < columnCount; i++)
                {
                    footerCells.Add(this.builder.InsertCell());
                    this.builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                }

                if (!table.Footer.ShowBorder)
                {
                    this.builder.RowFormat.Borders.Left.LineStyle = LineStyle.None;
                    this.builder.RowFormat.Borders.Right.LineStyle = LineStyle.None;
                    this.builder.RowFormat.Borders.Bottom.LineStyle = LineStyle.None;
                }

                this.builder.EndRow();

                this.builder.RowFormat.Borders.Left.LineStyle = oldLeftLineStyle;
                this.builder.RowFormat.Borders.Right.LineStyle = oldRightLineStyle;
                this.builder.RowFormat.Borders.Bottom.LineStyle = oldBottomLineStyle;
            }
            #endregion

            this.builder.EndTable();

            tb.AllowAutoFit = table.AllowAutoFit;
            tb.LeftIndent = leftIndent;

            this.InitFont();
        }

        private void AddTableCell(WordTableColumn column, object value, int rowIndex, Dictionary<int, Dictionary<string, object>> dictRows)
        {
            if (column.IsRowHeader)
            {
                this.CreateTableCell(column, true, value);
            }
            else
            {
                if (!WordUtil.IsEmptyColor(column.ValueFontColor))
                {
                    column.FontColor = column.ValueFontColor;
                }

                if (!column.AutoMerge)
                {
                    this.CreateTableValueCell(column, value);
                }
                else
                {
                    Dictionary<string, object> previousRow = rowIndex > 0 ? dictRows[rowIndex+1] : null;
                    string previousValue = previousRow == null ? null : (previousRow.ContainsKey(column.Name) ? previousRow[column.Name]?.ToString() : string.Empty);

                    this.CreateTableValueCell(column, value);
                    if (previousValue == value?.ToString())
                    {
                        this.builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                    else
                    {
                        this.builder.CellFormat.VerticalMerge = CellMerge.First;
                    }
                }
            }
        }

        /// <summary>
        /// 创建表格列头。
        /// </summary>
        /// <param name="column"></param>
        public void CreateTableHeaderCell(WordTableColumn column)
        {
            this.CreateTableCell(column, true, column.Title ?? column.Name);
        }

        /// <summary>
        /// 创建表格内容单元格。
        /// </summary>
        /// <param name="column"></param>
        /// <param name="value"></param>
        public void CreateTableValueCell(WordTableColumn column, object value)
        {
            this.CreateTableCell(column, false, value);
        }

        /// <summary>
        /// 创建表格单元格。
        /// </summary>
        /// <param name="column"></param>
        /// <param name="isHeader"></param>
        /// <param name="value"></param>
        private void CreateTableCell(WordTableColumn column, bool isHeader, object value)
        {
            this.builder.InsertCell();
            this.builder.Font.Size = column.FontSize;
            this.builder.Font.Color = column.FontColor;
            this.builder.CellFormat.HorizontalMerge = CellMerge.None;
            this.builder.CellFormat.VerticalMerge = CellMerge.None;
            this.builder.CellFormat.VerticalAlignment = column.VerticalAlignment;
            this.builder.ParagraphFormat.Alignment = column.Alignment;
            this.builder.CellFormat.Width = column.ActualWidth;        

            Color backgroundColor = isHeader ? WordUtil.PickColor(column.BackgroundColor, this.option.TableHeaderCellBackgroundColor) : column.BackgroundColor;

            this.builder.CellFormat.Shading.BackgroundPatternColor = backgroundColor;

            string strValue = value?.ToString()??"";

            if(isHeader)
            {
                this.builder.Write(strValue);
            }
            else
            {
                if(value!=null)
                {
                    #region Image
                    if (value is WordImage || value is IEnumerable<WordImage>)
                    {
                        List<WordImage> images = value is WordImage ? new List<WordImage>() { value as WordImage } : (value as IEnumerable<WordImage>).ToList();

                        images.ForEach(item =>
                        {
                            Image img = item.GetImage();
                            if (img.Width > column.ActualWidth)
                            {
                                item.Image = ImageHelper.ResizeImage(img, (int)column.ActualWidth, img.Height);
                                item.Type = WordImageType.Image;
                            }

                            this.AddImage(item);
                        });
                    }
                    #endregion
                    #region OleObject
                    else if (value is WordOleObject || value is IEnumerable<WordOleObject>)
                    {
                        List<WordOleObject> oleObjects = value is WordOleObject ? new List<WordOleObject>() { value as WordOleObject } : (value as IEnumerable<WordOleObject>).ToList();
                        int i = 0;
                        oleObjects.ForEach(item =>
                        {
                            this.AddOleObject(item, i < oleObjects.Count - 1);
                            i++;
                        });
                    } 
                    #endregion
                    else
                    {
                        this.builder.Write(strValue);
                    }
                }
                else
                {
                    this.builder.Write(strValue);
                }
            }           
        }
        #endregion    
     
        public void AddImage(WordImage img)
        {
            this.ClearParagraphFormat();          

            this.builder.Writeln();
            
            Image image = null;

            switch (img.Type)
            {              
                case WordImageType.FilePath:           
                    if (File.Exists(img.FilePath))
                    {
                        image = Image.FromFile(img.FilePath);                 
                    }
                    break;
                case WordImageType.Stream:
                    if (img.Stream != null)
                    {
                        image = Image.FromStream(img.Stream);              
                    }
                    break;
                case WordImageType.Bytes:
                    if (img.Bytes != null)
                    {
                        image = Image.FromStream(new MemoryStream(img.Bytes));
                    }
                    break;
                case WordImageType.Image:
                    if (img.Image != null)
                    {
                        image = img.Image;
                    }
                    break;
            }

            if (image!=null)
            {
                int maxWidth = (int)(this.option.PageWidth - this.builder.PageSetup.LeftMargin);
                if(img.FitPageWidth && image.Width> maxWidth)
                {
                    image = ImageHelper.ResizeImage(image, maxWidth, image.Height);
                }

                this.builder.InsertImage(image);

                if(img.Type==WordImageType.FilePath  && img.DeleteFileAfterUsed && File.Exists(img.FilePath))
                {
                    File.Delete(img.FilePath);
                }

                this.builder.Writeln();
            }
        }
       
        public void AddOleObject(WordOleObject obj, bool insertNewLine=true)
        {
            this.ClearParagraphFormat();

            this.builder.ParagraphFormat.LeftIndent = obj.LeftIndent;

            switch (obj.Type)
            {
                case WordOleObjectType.FilePath:
                    if (!string.IsNullOrEmpty(obj.FilePath))
                    {
                        Image icon = obj.IconImage;

                        if (obj.ShowAsIcon && (obj.ShowAsDetailList || obj.IconAsThumbnail))
                        {
                            icon = ImageHelper.ResizeImage(icon, obj.ThumbnailIconWidth, obj.ThumbnailIconHeight);
                        }
                       
                        this.builder.InsertOleObject(obj.FilePath, false, obj.ShowAsIcon, icon);

                        if(!obj.ShowAsDetailList && obj.IconPosition==WordOleObjectIconPosition.Top)
                        {
                            this.builder.Writeln();
                        }                     
                        
                        this.builder.Write(obj.Name);

                        if(insertNewLine)
                        {
                            this.builder.Writeln("");
                        }

                        if (obj.DeleteFileAfterUsed && File.Exists(obj.FilePath))
                        {
                            File.Delete(obj.FilePath);
                        }
                    }
                    break;
                case WordOleObjectType.Stream:
                    if (obj.Stream != null)
                    {
                        if (string.IsNullOrEmpty(obj.Id))
                        {
                            obj.Id = Guid.NewGuid().ToString();
                        }

                        this.builder.InsertOleObject(obj.Stream, obj.Id, obj.ShowAsIcon, obj.IconImage);
                        obj.Stream.Close();
                        this.builder.Writeln();
                    }

                    this.builder.Writeln(obj.Name);
                    break;
            }
        }
    }
}
