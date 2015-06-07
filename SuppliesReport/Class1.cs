using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SuppliesReport
{
    class Class1
    {
        /**
        public byte[] GenerateZayavlenie(ReportData data)
        {
            string filePath = HttpContext.Current.Server.MapPath("~/App_Content/RFC/zayavlenie.docx");
            byte[] templateData = System.IO.File.ReadAllBytes(filePath);

            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(templateData, 0, templateData.Length);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    SdtBlock overtimeBlock = GetContentBlockByTag(doc, "idOvertimeTable");
                    FillTimeTable(overtimeBlock, data.OvertimeTimes);

                    SdtBlock holidaysTimeBlock = GetContentBlockByTag(doc, "idHolidaysTimeTable");
                    FillTimeTable(holidaysTimeBlock, data.WeekendTimes);

                    SdtBlock nightTimeBlock = GetContentBlockByTag(doc, "idNightTimeTable");
                    FillTimeTable(nightTimeBlock, data.NightWorkTimes);

                    SdtBlock restTimeBlock = GetContentBlockByTag(doc, "idRestTimeTable");
                    FillTimeTable(restTimeBlock, data.RestTimes);


                    SdtBlock signaturesBlock = GetContentBlockByTag(doc, "idSignatures");

                    var names = data.RestTimes.Select(e => e.Name).Distinct().OrderBy(e => e);
                    foreach (string name in names)
                    {
                        string signatureLine = String.Format(
                            SignatureFormat,
                            DateTimeConverter.ToFormatedString(data.GenerationTime, DateTimeConverter.DefaultDateFormat),
                            name);

                        signaturesBlock.Append(new Paragraph(new Run(new Text(signatureLine))));
                    }
                }

                return mem.ToArray();
            }
        }

        private SdtBlock GetContentBlockByTag(WordprocessingDocument doc, string tagId)
        {
            SdtBlock block = doc.MainDocumentPart.Document.Body.Descendants<SdtBlock>()
                .Where(e => e.SdtProperties.GetFirstChild<Tag>().Val == tagId)
                .SingleOrDefault();

            return block;
        }

        private void FillPositionTimeTable(SdtBlock block, IEnumerable<ReportTimeItem> times)
        {
            if (block != null)
            {
                Table table = block.SdtContentBlock.GetFirstChild<Table>();

                var orderedTimes = times.OrderBy(e => e.WorkStartTime).ThenBy(e => e.Name);
                foreach (ReportTimeItem time in orderedTimes)
                {
                    TableRow row = table.AppendChild<TableRow>(new TableRow());

                    TableCell positionCell = row.AppendChild<TableCell>(new TableCell());
                    positionCell.Append(new Paragraph(new Run(new Text(time.Position))));

                    TableCell nameCell = row.AppendChild<TableCell>(new TableCell());
                    nameCell.Append(new Paragraph(new Run(new Text(time.Name))));

                    TableCell workStartCell = row.AppendChild<TableCell>(new TableCell());
                    string workStartValue = DateTimeConverter.ToFormatedString(time.WorkStartTime.Date, DateTimeConverter.DefaultDateFormat);
                    workStartCell.Append(new Paragraph(new Run(new Text(workStartValue))));

                    TableCell startCell = row.AppendChild<TableCell>(new TableCell());
                    string startValue = DateTimeConverter.ToFormatedString(time.StartTime, DateTimeConverter.FullDateTimeFormat);
                    startCell.Append(new Paragraph(new Run(new Text(startValue))));

                    TableCell endCell = row.AppendChild<TableCell>(new TableCell());
                    string endValue = DateTimeConverter.ToFormatedString(time.EndTime, DateTimeConverter.FullDateTimeFormat);
                    endCell.Append(new Paragraph(new Run(new Text(endValue))));
                }
            }
        }

        /**/
    }
}
