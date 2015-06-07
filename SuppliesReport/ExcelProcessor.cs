using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace SuppliesReport
{
    public class ExcelProcessor
    {
        SharedStringTable sharedStrings;
        IEnumerable<Sheet> workSheets;


        public List<List<PetitionGeneral>> Boo(string filePath)
        {
            //const string filePath = @"C:\Users\YPV.OSP\Documents\";
            //const string inputFileName = @"1 Petition draft input.xlsx";

            Workbook workBook;
            //WorksheetPart custSheet;

            List<List<PetitionGeneral>> petitions = new List<List<PetitionGeneral>>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                workBook = document.WorkbookPart.Workbook;
                workSheets = workBook.Descendants<Sheet>();
                sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                foreach (var sheet in workSheets)
                {
                    WorksheetPart wshp = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                    List<PetitionGeneral> pts = LoadPetitions(wshp.Worksheet, sharedStrings);
                    petitions.Add(pts);
                }
                //custID = workSheets.First().Id;
                //custSheet = (WorksheetPart)document.WorkbookPart.GetPartById(custID);

                //petitions = LoadCustomers(custSheet.Worksheet, sharedStrings);

            }

            return petitions;
        }

        public static List<PetitionGeneral> LoadPetitions(Worksheet worksheet, SharedStringTable sharedString)
        {
            List<PetitionGeneral> result = new List<PetitionGeneral>();

            IEnumerable<Row> dataRows = worksheet.Descendants<Row>();
            //int c = dataRows.Count();
            //dataRows = dataRows.Where(r => r.RowIndex > 0 && r.RowIndex < c - 1);

            foreach (Row row in dataRows)
            {
                IEnumerable<String> textValues = row.Descendants<Cell>()
                    .Where(cell => cell.CellValue != null)
                    .Select(cell =>
                        cell.DataType != null
                        && cell.DataType.HasValue
                        && cell.DataType == CellValues.SharedString
                        ? sharedString.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText
                        : cell.CellValue.InnerText);

                //Check to verify the row contained data.
                if (textValues.Count() > 0)
                {
                    //Create a customer and add it to the list.
                    var textArray = textValues.ToArray();
                    double d = double.Parse(textArray[6]);
                    PetitionGeneral customer = new PetitionGeneral()
                    {
                        ConsecutiveNumber = textArray[0],
                        Name = textArray[1],
                        Cost = textArray[2],
                        InventoryNumber = textArray[3],
                        Count = textArray[4],
                        ManufacuringYear = textArray[5],
                        AcceptedDate = DateTime.FromOADate(d),
                        LifeTime = textArray[7],
                        ActualPeriod = textArray[8]
                    };
                    result.Add(customer);
                }
                else
                {
                    //If no cells, then you have reached the end of the table.
                    break;
                }
            }

            //Return populated list of customers.
            return result;
        }
    }
}
