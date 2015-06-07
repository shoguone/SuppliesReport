using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SuppliesReport.Utility;
using SuppliesReport.EntityModel;

namespace SuppliesReport
{
    public class WordProcessor
    {
        private const string centerStyleId = "myCenteredStyleID";
        private const string leftStyleId = "myLeftStyleID";

        private string templatesDirectoryPath = Properties.Settings.Default.TemplatesDirectoryPath;
        private string templateReportFileName = Properties.Settings.Default.TemplateReportFileName;
        private string templateSupplyPopFileName = Properties.Settings.Default.TemplateSupplyPopFileName;
        private string templateSupplyPushFileName = Properties.Settings.Default.TemplateSupplyPushFileName;
        
        private string templateActInspectionFileName = Properties.Settings.Default.TemplateActInspectionFileName;
        private string templateActWritingOffFileName = Properties.Settings.Default.TemplateActWritingOffFileName;
        private string templateActEliminationFileName = Properties.Settings.Default.TemplateActEliminationFileName;
        private string templateProtocolSessionFileName = Properties.Settings.Default.TemplateProtocolSessionFileName;
        private string templateOrderWritingOffFileName = Properties.Settings.Default.TemplateOrderWritingOffFileName;

        private string outputDirectoryPath = Properties.Settings.Default.OutputDirectoryPath;

        private string writingoffInventorynameHelpString = Properties.Settings.Default.WritingoffInventorynameHelpString;


        public void FillSupplyPop(DateTime firingDate)
        {
            string fname = templateSupplyPopFileName;
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + fname);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    IncludeStyles(doc);

                    List<PRIZ> data;
                    using (var ctx = new FormContext())
                    {
                        data = ctx.PRIZ
                            .Where(p => p.FL_UV == 1 && p.D_U_UVOL.HasValue)
                            .OrderBy(p => p.FAM).ThenBy(p => p.IM).ThenBy(p => p.OTCH)
                            .AsEnumerable()
                            .Where(p => p.D_U_UVOL.Value.Date == firingDate)
                            .ToList();
                    }
                    SdtBlock list = GetContentBlockByTag(doc, "tagList");
                    FillSupplyPopTable(list, data);

                    SdtRun dateHeader = GetContentRunByTag(doc, "tagDateHeader");
                    SetFieldText(dateHeader, RuDateAndMoneyConverter.DateToTextLong(firingDate, "г."));

                    SdtRun recruitsCount = GetContentRunByTag(doc, "tagRecruitsCount");
                    SetFieldText(recruitsCount, data.Count.ToString());

                    int pagesCount = 0;
                    if (int.TryParse(doc.ExtendedFilePropertiesPart.Properties.Pages.Text, out pagesCount))
                        pagesCount -= 1;
                    SdtRun pagesCountRun = GetContentRunByTag(doc, "tagPagesCount");
                    SetFieldText(pagesCountRun, pagesCount.ToString());

                    SdtRun dateFooter = GetContentRunByTag(doc, "tagDateFooter");
                    SetFieldText(dateFooter, firingDate.ToShortDateString());

                    //var sum = data.Sum(pt => pt.SummaryCost);
                    //SdtBlock overall = GetContentBlockByTag(doc, "tagOverall");
                    //SetFieldText(overall, sum.ToString("F2"));
                }

                int fnLength = fname.Length;
                int extLength = 5;
                int templateLength = 7;
                string outputFileName = fname.Substring(templateLength, fnLength - templateLength - extLength);
                string extension = fname.Substring(fnLength - extLength, extLength);
                File.WriteAllBytes(string.Format("{0}{1} {2}{3}", outputDirectoryPath, outputFileName, firingDate.ToShortDateString(), extension), mem.ToArray());
            }
        }

        public void FillSupplyPush(DateTime hiringDate)
        {
            string fname = templateSupplyPushFileName;
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + fname);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    IncludeStyles(doc);

                    List<PRIZ> data;
                    using (var ctx = new FormContext())
                    {
                        data = ctx.PRIZ
                            .Where(p => p.FL_UV == 1 && p.D_U_UVOL.HasValue)
                            .OrderBy(p => p.FAM).ThenBy(p => p.IM).ThenBy(p => p.OTCH)
                            .AsEnumerable()
                            .Where(p => p.D_P_UVOL.Value.Date == hiringDate)
                            .ToList();
                    }
                    SdtBlock list = GetContentBlockByTag(doc, "tagList");
                    FillSupplyPopTable(list, data);

                    SdtRun dateHeader = GetContentRunByTag(doc, "tagDateHeader");
                    SetFieldText(dateHeader, RuDateAndMoneyConverter.DateToTextLong(hiringDate, "г."));

                    SdtRun recruitsCount = GetContentRunByTag(doc, "tagRecruitsCount");
                    SetFieldText(recruitsCount, data.Count.ToString());

                    int pagesCount = 0;
                    if (int.TryParse(doc.ExtendedFilePropertiesPart.Properties.Pages.Text, out pagesCount))
                        pagesCount -= 1;
                    SdtRun pagesCountRun = GetContentRunByTag(doc, "tagPagesCount");
                    SetFieldText(pagesCountRun, pagesCount.ToString());

                    SdtRun dateFooter = GetContentRunByTag(doc, "tagDateFooter");
                    SetFieldText(dateFooter, hiringDate.ToShortDateString());
                }

                int fnLength = fname.Length;
                int extLength = 5;
                int templateLength = 7;
                string outputFileName = fname.Substring(templateLength, fnLength - templateLength - extLength);
                string extension = fname.Substring(fnLength - extLength, extLength);
                File.WriteAllBytes(string.Format("{0}{1} {2}{3}", outputDirectoryPath, outputFileName, hiringDate.ToShortDateString(), extension), mem.ToArray());
            }
        }

        #region FILL methods for filling acts templates with input data
        public void FillActInspection(IEnumerable<PetitionGeneral> data, int consNum)
        {
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + templateActInspectionFileName);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    IncludeStyles(doc);

                    SdtBlock testingNInspection = GetContentBlockByTag(doc, "tagTestingNInspection");
                    FillActInspectionTable(testingNInspection, data);

                    var sum = data.Sum(pt => pt.SummaryCost);
                    SdtBlock overall = GetContentBlockByTag(doc, "tagOverall");
                    SetFieldText(overall, sum.ToString("F2"));

                    SdtRun overallFooter = GetContentRunByTag(doc, "tagOverallFooter");
                    SetFieldText(overallFooter, RuDateAndMoneyConverter.CurrencyToTxtShorter(sum));

                    SdtRun overallString = GetContentRunByTag(doc, "tagOverallString");
                    SetFieldText(overallString, RuDateAndMoneyConverter.CurrencyToTxtWithCopecks(sum, true));
                }

                string outputFileName = templateActInspectionFileName.Substring(7);
                File.WriteAllBytes(string.Format("{0}{1} {2}", outputDirectoryPath, consNum, outputFileName), mem.ToArray());
                //return mem.ToArray();
            }
        }

        public void FillActWritingOff(IEnumerable<PetitionGeneral> data, int consNum)
        {
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + templateActWritingOffFileName);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    //IncludeStyles(doc);

                    SdtBlock writingOff = GetContentBlockByTag(doc, "tagWritingOff");
                    FillActWritingOffTable(writingOff, data);

                    int countSum = (int)data.Sum(pt => pt.CountFloat);
                    SdtBlock count = GetContentBlockByTag(doc, "tagCountTable");
                    SetFieldText(count, countSum.ToString());

                    float sum = data.Sum(pt => pt.SummaryCost);
                    SdtBlock overall = GetContentBlockByTag(doc, "tagOverallTable");
                    SetFieldText(overall, sum.ToString("F2"));

                    SdtBlock countString = GetContentBlockByTag(doc, "tagCountString");
                    SetFieldText(countString, RuDateAndMoneyConverter.NumeralsToTxt((long)countSum, TextCase.Nominative, true, true));

                    SdtBlock overallFooter = GetContentBlockByTag(doc, "tagOverallFooter");
                    SetFieldText(overallFooter, sum.ToString("F2"));

                    SdtBlock overallString = GetContentBlockByTag(doc, "tagOverallString");
                    SetFieldText(overallString, RuDateAndMoneyConverter.CurrencyToTxtWithCopecks(sum, true));
                }

                string outputFileName = templateActWritingOffFileName.Substring(7);
                File.WriteAllBytes(string.Format("{0}{1} {2}", outputDirectoryPath, consNum, outputFileName), mem.ToArray());
            }
        }

        public void FillActElimination(PetitionGeneral data, int listNum, int itemNum)
        {
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + templateActEliminationFileName);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    SdtRun name = GetContentRunByTag(doc, "tagName");
                    SetFieldText(name, data.Name);

                    SdtBlock invNum = GetContentBlockByTag(doc, "tagInventoryNumber");
                    SetFieldText(invNum, data.InventoryNumber);
                    
                    SdtRun name2 = GetContentRunByTag(doc, "tagName2");
                    SetFieldText(name2, data.Name);

                }

                string outputFileName = templateActEliminationFileName.Substring(7);
                File.WriteAllBytes(string.Format("{0}{1}_{2} {3}", outputDirectoryPath, listNum, itemNum, outputFileName), mem.ToArray());
            }
        }

        public void FillProtocolSession(IEnumerable<PetitionGeneral> data, int consNum)
        {
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + templateProtocolSessionFileName);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    SdtBlock list = GetContentBlockByTag(doc, "tagList");
                    FillProtocolSessionList(list, data);

                    int countSum = (int)data.Sum(pt => pt.CountFloat);
                    SdtRun count = GetContentRunByTag(doc, "tagCount");
                    string result = string.Format("{0} {1}", countSum, "единиц");
                    if (countSum % 10 == 1)
                        result += "ы";
                    SetFieldText(count, result);
                }

                string outputFileName = templateProtocolSessionFileName.Substring(7);
                File.WriteAllBytes(string.Format("{0}{1} {2}", outputDirectoryPath, consNum, outputFileName), mem.ToArray());
            }
        }

        public void FillOrder(IEnumerable<PetitionGeneral> data, int consNum)
        {
            byte[] templateData = File.ReadAllBytes(templatesDirectoryPath + templateOrderWritingOffFileName);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                CustomFormat format = new CustomFormat(JustificationValues.Left, 14) { LineSpacingAfter = "120" };

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    SdtBlock list = GetContentBlockByTag(doc, "tagList");
                    FillOrderList(list, data, format);

                    float sum = data.Sum(pt => pt.SummaryCost);
                    SdtRun overall = GetContentRunByTag(doc, "tagOverall");
                    SetFieldText(overall, sum.ToString("F2"));

                    SdtRun overallString = GetContentRunByTag(doc, "tagOverallString");
                    SetFieldText(overallString, RuDateAndMoneyConverter.CurrencyToTxtWithCopecks(sum, true));
                }

                string outputFileName = templateOrderWritingOffFileName.Substring(7);
                File.WriteAllBytes(string.Format("{0}{1} {2}", outputDirectoryPath, consNum, outputFileName), mem.ToArray());
            }
        }
        #endregion FILL

        /// <summary>
        /// First testing method
        /// </summary>
        /// <param name="data"></param>
        public void Foo(List<PetitionGeneral> data)
        {
            const string filePath = @"C:\Users\YPV.OSP\Documents\";
            const string inputFileName = @"1 act input.docx";
            const string outputFileName = @"1 act output.docx";
            byte[] templateData = File.ReadAllBytes(filePath + inputFileName);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, false))
                {
                    IncludeStyles(doc);

                    SdtBlock testingNInspection = GetContentBlockByTag(doc, "tagTestingNInspection");
                    FillActInspectionTable(testingNInspection, data);

                    var sum = data.Sum(pt => pt.SummaryCost);
                    SdtBlock overall = GetContentBlockByTag(doc, "tagOverall");
                    SetFieldText(overall, sum.ToString("F2"));

                }

                File.WriteAllBytes(filePath + outputFileName, mem.ToArray());
                //return mem.ToArray();
            }

        }

        private void FillSupplyPopTable(SdtBlock block, List<PRIZ> entities)
        {
            if (block != null)
            {
                Table table = block.SdtContentBlock.GetFirstChild<Table>();
                //var lastRow = table.Descendants<TableRow>().LastOrDefault();

                int i = 1;
                foreach (PRIZ entity in entities)
                {
                    TableRow row = table.AppendChild<TableRow>(new TableRow());
                    //TableRow row = table.InsertBefore<TableRow>(new TableRow(), lastRow);

                    AddStyledTextToTheRow(row, i.ToString(), centerStyleId);
                    AddStyledTextToTheRow(row, string.Format("{0} {1} {2}",
                        entity.FAM, entity.IM, entity.OTCH), leftStyleId);
                    AddStyledTextToTheRow(row, entity.D_ROD.GetValueOrDefault().ToShortDateString(), centerStyleId);
                    AddStyledTextToTheRow(row, entity.RVK, leftStyleId);
                    //AddStyledTextToTheRow(row, entity.CostFloat.ToString("F2"), centerStyleId);
                    //AddStyledTextToTheRow(row, entity.SummaryCost.ToString("F2"), centerStyleId);
                    i++;
                }
            }
        }

        #region old methods
        //private void FillTable(SdtBlock block, IEnumerable<ActOfTestingAndInspection> entities)
        private void FillActInspectionTable(SdtBlock block, IEnumerable<PetitionGeneral> entities)
        {
            if (block != null)
            {
                Table table = block.SdtContentBlock.GetFirstChild<Table>();
                var lastRow = table.Descendants<TableRow>().LastOrDefault();

                foreach (PetitionGeneral entity in entities)
                {
                    //TableRow row = table.AppendChild<TableRow>(new TableRow());
                    TableRow row = table.InsertBefore<TableRow>(new TableRow(), lastRow);

                    AddStyledTextToTheRow(row, entity.ConsecutiveNumber, centerStyleId);
                    AddStyledTextToTheRow(row, entity.Name, leftStyleId);
                    AddStyledTextToTheRow(row, entity.InventoryNumber, centerStyleId);
                    AddStyledTextToTheRow(row, entity.Count, centerStyleId);
                    AddStyledTextToTheRow(row, entity.CostFloat.ToString("F2"), centerStyleId);
                    AddStyledTextToTheRow(row, entity.SummaryCost.ToString("F2"), centerStyleId);

                }
            }
        }

        private void FillActWritingOffTable(SdtBlock block, IEnumerable<PetitionGeneral> entities)
        {
            if (block != null)
            {
                Table table = block.SdtContentBlock.GetFirstChild<Table>();
                var lastRow = table.Descendants<TableRow>().LastOrDefault();

                foreach (PetitionGeneral entity in entities)
                {
                    TableRow row = table.InsertBefore<TableRow>(new TableRow(), lastRow);

                    CustomFormat formatLeft = new CustomFormat(JustificationValues.Left, 9);
                    CustomFormat formatCenter = new CustomFormat(JustificationValues.Center, 9);

                    AddCellWithParametersToTheRow(row, entity.ConsecutiveNumber, formatCenter);
                    AddCellWithParametersToTheRow(row, entity.Name +
                        writingoffInventorynameHelpString +
                        entity.InventoryNumber, formatLeft);
                    AddCellWithParametersToTheRow(row, "", formatCenter);
                    AddCellWithParametersToTheRow(row, entity.Count, formatCenter);
                    AddCellWithParametersToTheRow(row, entity.CostFloat.ToString("F2"), formatCenter);
                    AddCellWithParametersToTheRow(row, entity.SummaryCost.ToString("F2"), formatCenter);
                    AddCellWithParametersToTheRow(row, "", formatLeft);
                    AddCellWithParametersToTheRow(row, "", formatLeft);

                }
            }
        }

        private void FillProtocolSessionList(SdtBlock block, IEnumerable<PetitionGeneral> entities)
        {
            if (block != null)
            {
                var np = (block.SdtContentBlock.FirstChild as Paragraph).ParagraphProperties.NumberingProperties;
                block.SdtContentBlock.RemoveAllChildren<Paragraph>();

                var entitiesGrouped = entities.GroupBy(e => e.Name);
                int i = 0, c = entitiesGrouped.Count();

                foreach (var entity in entitiesGrouped)
                {
                    i++;
                    string listItem = string.Format("{0} - {1} ед.{2}", 
                        entity.Key, 
                        entity.Count(), 
                        i != c ? ";" : string.Empty);
                    Run r = new Run(new Text(listItem));
                    Paragraph p = new Paragraph(r)
                    {
                        ParagraphProperties = new ParagraphProperties()
                        {
                            NumberingProperties = new NumberingProperties()
                            {
                                NumberingId = new NumberingId() { Val = np.NumberingId.Val },
                                NumberingLevelReference = new NumberingLevelReference() { Val = np.NumberingLevelReference.Val }
                            }
                        }
                    };

                    block.SdtContentBlock.AppendChild<Paragraph>(p);

                }
            }
        }

        private void FillOrderList(SdtBlock block, IEnumerable<PetitionGeneral> entities,
            CustomFormat format)
        {
            FillOrderList(block, entities, format.FontSizeString, format.Font, format.JustificationValue,
                format.LineSpacing, format.LineSpacingRule, format.LineSpacingBefore, format.LineSpacingAfter);
        }

        private void FillOrderList(SdtBlock block, IEnumerable<PetitionGeneral> entities,
            string fontSizeString, string fontString,
            JustificationValues justificationValue, string lineSpacing, LineSpacingRuleValues lineSpacingRule, 
            string lineSpacingBefore, string lineSpacingAfter)
        {
            if (block != null)
            {
                Paragraph firstParagraph = block.SdtContentBlock.GetFirstChild<Paragraph>();
                var np = firstParagraph.ParagraphProperties.NumberingProperties;
                block.SdtContentBlock.RemoveAllChildren<Paragraph>();

                var entitiesGrouped = entities.GroupBy(e => e.Name);
                int i = 0, 
                    c = entitiesGrouped.Count();

                foreach (var entity in entitiesGrouped)
                {
                    i++;

                    string items = string.Empty,
                        years,
                        year,
                        cost;
                    int i1 = 0,
                        c1 = entity.Count();
                    foreach (var e in entity)
                    {
                        i1++;

                        items += e.InventoryNumber;
                        if (i1 < c1)
                            items += "; ";
                    }

                    var petition = entity.FirstOrDefault();
                    years = (DateTime.Now.Year - petition.AcceptedDate.Year).ToString();
                    year = petition.AcceptedDate.Year.ToString();
                    cost = petition.Cost.Replace('.', ',');

                    string listItem = string.Format("{0} (инв. №№ {1}), лет в эксплуатации - {2}, год ввода в эксплуатацию - {3}, цена за 1 шт. - {4} р.",
                        entity.Key,
                        items,
                        years,
                        year,
                        cost);
                    if (i < c)
                        listItem += ";";

                    FontSize fontSize = new FontSize() { Val = fontSizeString };
                    RunFonts font = new RunFonts()
                    {
                        Ascii = fontString,
                        EastAsia = fontString,
                        HighAnsi = fontString,
                        ComplexScript = fontString
                    };
                    Justification justification = new Justification() { Val = justificationValue };
                    SpacingBetweenLines spacingBwLines = new SpacingBetweenLines()
                    {
                        Line = lineSpacing,
                        LineRule = lineSpacingRule,
                        Before = lineSpacingBefore,
                        After = lineSpacingAfter
                    };


                    Run r = new Run(new Text(listItem))
                    {
                        RunProperties = new RunProperties()
                        {
                            FontSize = fontSize,
                            RunFonts = font
                        }
                    };
                    Paragraph p = new Paragraph(r)
                    {
                        ParagraphProperties = new ParagraphProperties()
                        {
                            NumberingProperties = new NumberingProperties()
                            {
                                NumberingId = new NumberingId() { Val = np.NumberingId.Val },
                                NumberingLevelReference = new NumberingLevelReference() { Val = np.NumberingLevelReference.Val }
                            },
                            Justification = justification,
                            SpacingBetweenLines = spacingBwLines
                        }
                    };

                    block.SdtContentBlock.AppendChild<Paragraph>(p);

                }
            }
        }
        #endregion old methods
        private void SetFieldText(SdtBlock block, string text)
        {
            var p = block.SdtContentBlock.GetFirstChild<Paragraph>();
            var r = p.GetFirstChild<Run>();
            var t = r.GetFirstChild<Text>();

            t.Text = text;

        }

        private void SetFieldText(SdtRun run, string text)
        {
            var r = run.SdtContentRun.GetFirstChild<Run>();
            var t = r.GetFirstChild<Text>();

            t.Text = text;

        }

        #region GET different types of sdt for further filling
        private SdtElement GetContentByTag(WordprocessingDocument doc, string tagId)
        {
            SdtBlock block = doc.MainDocumentPart.Document.Body.Descendants<SdtBlock>()
                .Where(e => e.SdtProperties.GetFirstChild<Tag>().Val == tagId)
                .SingleOrDefault();
            if (block == null)
            {
                SdtRun run = doc.MainDocumentPart.Document.Body.Descendants<SdtRun>()
                    .Where(e => e.SdtProperties.GetFirstChild<Tag>().Val == tagId)
                    .SingleOrDefault();
                return run;
            }
            return block;
        }

        private SdtBlock GetContentBlockByTag(WordprocessingDocument doc, string tagId)
        {
            //var az = doc.MainDocumentPart.Document.Body.Descendants<Tag>();
            //var a = doc.MainDocumentPart.Document.Body.Descendants<SdtBlock>();
            //var b = a.Where(e => e.SdtProperties.GetFirstChild<Tag>().Val == tagId);
            //var c = b.SingleOrDefault();
            SdtBlock block = doc.MainDocumentPart.Document.Body.Descendants<SdtBlock>()
                .Where(e => e.SdtProperties.GetFirstChild<Tag>().Val == tagId)
                .SingleOrDefault();

            return block;
        }

        private SdtRun GetContentRunByTag(WordprocessingDocument doc, string tagId)
        {
            SdtRun run = doc.MainDocumentPart.Document.Body.Descendants<SdtRun>()
                .Where(e => e.SdtProperties.GetFirstChild<Tag>().Val == tagId)
                .SingleOrDefault();

            return run;
        }
        #endregion GET

        private void AddStyledTextToTheRow(TableRow row, string text, string style)
        {
            TableCell positionCell = row.AppendChild<TableCell>(new TableCell());
            positionCell.TableCellProperties = new TableCellProperties() { TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center } };

            Run r = new Run(new Text(text));
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = style };
            Paragraph p = new Paragraph(r) { ParagraphProperties = new ParagraphProperties() { ParagraphStyleId = paragraphStyleId } };
            positionCell.Append(p);
        }

        private void AddCellWithParametersToTheRow(TableRow row, string text, CustomFormat format)
        {
            AddCellWithParametersToTheRow(row, text, 
                format.CellWidthString, format.JustificationValue, format.TableVerticalAlignmentValue, 
                format.LineSpacing, format.LineSpacingRule, format.LineSpacingBefore, format.LineSpacingAfter,
                format.FontSizeString, format.Font);
        }

        private void AddCellWithParametersToTheRow(TableRow row, string text, 
            string width, 
            JustificationValues justificationValue, TableVerticalAlignmentValues tableVerticalAlignmentValue,
            string lineSpacing, LineSpacingRuleValues lineSpacingRule, string lineSpacingBefore, string lineSpacingAfter,
            string fontSizeString, string fontString)
        {
            TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = tableVerticalAlignmentValue };

            Justification justification = new Justification() { Val = justificationValue };
            SpacingBetweenLines spacingBwLines = new SpacingBetweenLines()
            {
                Line = lineSpacing,
                LineRule = lineSpacingRule,
                Before = lineSpacingBefore,
                After = lineSpacingAfter
            };

            FontSize fontSize = new FontSize() { Val = fontSizeString };
            RunFonts font = new RunFonts()
            {
                Ascii = fontString,
                EastAsia = fontString,
                HighAnsi = fontString,
                ComplexScript = fontString
            };
            
            TableCell positionCell = row.AppendChild<TableCell>(new TableCell());
            positionCell.TableCellProperties = new TableCellProperties() { TableCellVerticalAlignment = tableCellVerticalAlignment };
            if (!string.IsNullOrEmpty(width))
            {
                TableCellWidth tableCellWidth = new TableCellWidth() { Width = width };
                positionCell.TableCellProperties.TableCellWidth = tableCellWidth;
            }

            Run r = new Run(new Text(text))
            {
                RunProperties = new RunProperties()
                {
                    FontSize = fontSize,
                    RunFonts = font
                }
            };

            Paragraph p = new Paragraph(r)
            {
                ParagraphProperties = new ParagraphProperties()
                {
                    Justification = justification,
                    SpacingBetweenLines = spacingBwLines
                }
            };

            positionCell.Append(p);
        }

        #region STYLES
        private void IncludeStyles(WordprocessingDocument doc)
        {
            // Get the Styles part for this document.
            StyleDefinitionsPart part = doc.MainDocumentPart.StyleDefinitionsPart;

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage(doc);
            }

            AddStyle(doc, centerStyleId,
                "Normal", centerStyleId,
                "22", "Times New Roman",
                "240", LineSpacingRuleValues.Auto,
                JustificationValues.Center
                //, TableVerticalAlignmentValues.Center
                );
            AddStyle(doc, leftStyleId,
                "Normal", leftStyleId,
                "22", "Times New Roman",
                "240", LineSpacingRuleValues.Auto,
                JustificationValues.Left
                //, TableVerticalAlignmentValues.Center
                );

        }

        private void AddStyle(WordprocessingDocument doc, string styleId,
            string basedOnStyle, string nextParagraphStyle,
            string fontSizeString, string fontString,
            string lineSpacing, LineSpacingRuleValues lineSpacingRule,
            JustificationValues justificationValue
            //, TableVerticalAlignmentValues tableVerticalAlignmentValues
            )
        {
            // Create a new paragraph style and specify some of the properties.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            };

            StyleName styleName = new StyleName() { Val = style.StyleId };
            BasedOn basedOn = new BasedOn() { Val = basedOnStyle };
            NextParagraphStyle nextParagraph = new NextParagraphStyle() { Val = nextParagraphStyle };

            style.Append(styleName);
            style.Append(basedOn);
            style.Append(nextParagraph);

            StyleParagraphProperties styleParagraphProperties = new StyleParagraphProperties();
            Justification justification = new Justification() { Val = justificationValue };
            SpacingBetweenLines spacingBwLines = new SpacingBetweenLines()
            {
                Line = lineSpacing,
                LineRule = lineSpacingRule
            };
            styleParagraphProperties.Append(justification);
            styleParagraphProperties.Append(spacingBwLines);

            StyleRunProperties styleRunProperties = new StyleRunProperties();
            FontSize fontSize = new FontSize() { Val = fontSizeString };
            //RunFonts font = new RunFonts() { Ascii = fontString };
            RunFonts font = new RunFonts()
            {
                Ascii = fontString,
                EastAsia = fontString,
                HighAnsi = fontString,
                ComplexScript = fontString
            };
            styleRunProperties.Append(fontSize);
            styleRunProperties.Append(font);

            //StyleTableCellProperties styleTableCellProperties = new StyleTableCellProperties();
            //TableCellVerticalAlignment verticalAlignment = new TableCellVerticalAlignment() { Val = tableVerticalAlignmentValues };
            //styleTableCellProperties.Append(verticalAlignment);

            style.Append(styleParagraphProperties);
            style.Append(styleRunProperties);
            //style.Append(styleTableCellProperties);

            // Add the style to the styles part.
            doc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(style);
        }

        /// Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }
        #endregion STYLES
    }
}
