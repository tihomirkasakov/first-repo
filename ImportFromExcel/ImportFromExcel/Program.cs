using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Data;

namespace ImportFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //HSSFWorkbook excelFileReader = new HSSFWorkbook();
            //List<ImprotedReservation> result = new List<ImprotedReservation>();
            //using (FileStream file = new FileStream(@"D:\test.xlsx", FileMode.Open, FileAccess.Read))
            //{
            //    excelFileReader = new HSSFWorkbook(file);
            //}
            //ISheet sheet = excelFileReader.GetSheet("List1");
            //for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            //{
            //    int colIndex = 0;
            //    int cellsOnRow = sheet.GetRow(rowIndex).LastCellNum;
            //    List<string> info = new List<string>();

            //    for (int cells = 0; cells < cellsOnRow; cells++)
            //    {
            //        info.Add(sheet.GetRow(rowIndex).GetCell(cells).ToString());
            //    }

            //    if (sheet.GetRow(rowIndex) != null && info.Count == 9)
            //    {
            //        string clientBrand = info[colIndex++];
            //        string voucher = info[colIndex++];
            //        string passengerName = info[colIndex++].Split(new string[] { "<BR>" }, StringSplitOptions.RemoveEmptyEntries).First();
            //        var passengers = info[colIndex++].Split(new char[] { '(', '/', ')', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            //        int totalPassengers = int.Parse(passengers[0]);
            //        bool hasAdults = int.TryParse(passengers[1], out int adults);
            //        bool hasChildrens = int.TryParse(passengers[2], out int childrens);
            //        bool hasInfants = int.TryParse(passengers[3], out int infants);
            //        string pickupTime = info[colIndex++];
            //        string pickupFrom = info[colIndex++];
            //        string hotelArea = info[colIndex++];
            //        string dropTo = info[colIndex++];
            //        var flightInfo = info[colIndex++].Split(new string[] { "<br>" }, StringSplitOptions.RemoveEmptyEntries);
            //        string flightNumber = flightInfo[0];
            //        string departureTime = flightInfo[1];
            //        string arriveTime = flightInfo[2];

            //        ImprotedReservation reservation = new ImprotedReservation{clientBrand, voucher, passengerName, totalPassengers, adults, childrens, infants, pickupTime, pickupFrom, hotelArea, dropTo, flightNumber, departureTime, arriveTime);
            //        result.Add(reservation);
            //    }
            //}



            string fileName = @"d:\test.xlsx";
            DataTable dt = new DataTable();
            List<string> Headers = new List<string>();
            bool firstRowIsHeader = true;

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                //Read the first Sheets 
                Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                foreach (Row row in rows)
                {
                    //Read the first row as header
                    if (row.RowIndex.Value == 1)
                    {
                        var j = 1;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
                            Console.WriteLine(colunmName);
                            Headers.Add(colunmName);
                            dt.Columns.Add(colunmName);
                        }
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
                            i++;
                        }
                    }
                }
            }

            string GetCellValue(SpreadsheetDocument doc, Cell cell)
            {
                string value = cell.CellValue.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                }
                return value;
            }

                foreach (DataRow row in dt.Rows)
                {
                var test= row[0].ToString();
            }

            //using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            //{
            //    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
            //    {
            //        WorkbookPart workbookPart = doc.WorkbookPart;
            //        SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            //        SharedStringTable sst = sstpart.SharedStringTable;

            //        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            //        Worksheet sheet = worksheetPart.Worksheet;

            //        var rows = sheet.Descendants<Row>();
            //        int colIndex = 0;
            //Console.WriteLine($"Row count = {rows.LongCount()}");
            //Console.WriteLine($"Cell count = {cells.LongCount()}");


            //foreach (Row row in rows)
            //{
            //    var cells = row.Elements<Cell>();
            //    int fkOperatingClient = 0;
            //    int index = 1;
            //string codeClient = CellValue(cells, index);
            //if (codeClient.IsNullOrWhiteSpace())
            //{ // пропускат се редовете без код на клиент - че накрая има доста празни.
            //    continue;
            //}
            //if (lastCodeClient.IsNullOrWhiteSpace())
            //{
            //    lastCodeClient = codeClient;
            //    operatingClient = operatingClients.GetWithContract(codeClient);
            //    if (operatingClient == null)
            //    {
            //        ex.Throw("Не е намерен код на клиент, не е активен или няма договор за превоз", codeClient);
            //    }
            //    fkOperatingClient = operatingClient.ID;
            //}
            //else if (lastCodeClient != codeClient)
            //{
            //    ex.Throw("Във файла трябва да има само един код на клиент!");
            //}
            //index += 2;
            //string codeBrand = CellValue(cells, index++);
            //DateTime dtEvent = DateTime.FromOADate(CellValue(cells, index++).ToInt());
            //int hour = DateTime.FromOADate(CellValue(cells, index++).ToDouble()).Hour;
            //dtEvent = dtEvent.AddHours(hour);
            //int minutes = CellValue(cells, index++).ToInt();
            //dtEvent = dtEvent.AddMinutes(minutes);
            //string flight = CellValue(cells, index++);
            //bool? isDeparture = null;
            //switch (CellValue(cells, index))
            //{
            //    case "DEP":
            //        isDeparture = true;
            //        break;

            //    case "ARR":
            //        isDeparture = false;
            //        break;
            //}
            //index += 2;
            //string codeResortFrom = CellValue(cells, index);
            //index += 2;
            //string codeHotelFrom = CellValue(cells, index);
            //index += 2;
            //string addressFrom = CellValue(cells, index++);
            //string codeResortTo = CellValue(cells, index);
            //index += 2;
            //string codeHotelTo = CellValue(cells, index);
            //index += 2;
            //string addressTo = CellValue(cells, index++);
            //int adults = CellValue(cells, index++).ToInt();
            //int children = CellValue(cells, index++).ToInt();
            //int infants = CellValue(cells, index).ToInt();
            //index += 2;
            //string voucher = CellValue(cells, index++);
            //string code = CellValue(cells, index++);
            //string passenger = CellValue(cells, index++);
            //string notes = CellValue(cells, index++);
            //NomResort resortFrom = resorts.Get(q => q.Code.Equals(codeResortFrom)) ?? new NomResort();
            //int fkResortFrom = GetResort(resortFrom);
            //Hotel hotelFrom = GetHotel(resortFrom, codeHotelFrom);
            //int? fkHotelFrom = hotelFrom != null ? (int?)hotelFrom.ID : null;
            //NomResort resortTo = resorts.Get(q => q.Code.Equals(codeResortTo)) ?? new NomResort();
            //int fkResortTo = GetResort(resortTo);
            //Hotel hotelTo = GetHotel(resortTo, codeHotelTo);
            //int? fkHotelTo = hotelTo != null ? (int?)hotelTo.ID : null;
            //int? fkBrand = operatingClient == null ? null : (int?)operatingClient.ClientBrands.Where(q => q.IsActive && q.Code.Equals(codeBrand)).Select(q => (int?)q.ID).FirstOrDefault();
            //ClientBrand brand = fkBrand.HasValue ? operatingClient.ClientBrands.First(q => q.ID.Equals(fkBrand.Value)) : null;

            //Reservation reservation = new Reservation
            //{
            //    FKOperatingClient = fkOperatingClient,
            //    OperatingClient = operatingClient,
            //    FKBrand = fkBrand,
            //    ClientBrand = brand,
            //    FKNomResortFrom = fkResortFrom,
            //    NomResort1 = resortFrom,
            //    FKNomResortTo = fkResortTo,
            //    NomResort = resortTo,
            //    FKHotelFrom = fkHotelFrom,
            //    Hotel = hotelFrom,
            //    FKHotelTo = fkHotelTo,
            //    Hotel1 = hotelTo,
            //    Adults = adults,
            //    Children = children,
            //    Infants = infants,
            //    HotelAddressFrom = addressFrom,
            //    HotelAddressTo = addressTo,
            //    //Address = addressFrom.IsNullOrWhiteSpace() ? addressTo : addressFrom,
            //    FlightNumber = flight,
            //    Notes = notes,
            //    IsDeparture = isDeparture,
            //    dtEvent = dtEvent,
            //    Code = code,
            //    PassengerName = passenger,
            //    Reference = voucher,
            //};
            //result.Add(reservation);

            //}
            //    }
            //}

            //new OpenXmlExcel().ExcelToCsv(fileName, "f1.csv", ";", true);
        }
        //public class OpenXmlExcel
        //{
        //    public void ExcelToCsv(string source, string target, string delimiter = ";", bool firstRowIsHeade = true)
        //    {
        //        var dt = ReadExcelSheet(source, firstRowIsHeade);
        //        DatatableToCsv(dt, target, delimiter);

        //    }

        //    private void DatatableToCsv(DataTable dt, string fname, string delimiter = ";")
        //    {

        //        using (StreamWriter writer = new StreamWriter(fname))
        //        {
        //            foreach (DataRow row in dt.AsEnumerable())
        //            {
        //                writer.WriteLine(string.Join(delimiter, row.ItemArray.Select(x => x.ToString())) + delimiter);
        //            }
        //        }

        //    }

        //    List<string> Headers = new List<string>();


        //    private DataTable ReadExcelSheet(string fname, bool firstRowIsHeade)
        //    {

        //        DataTable dt = new DataTable();
        //        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fname, false))
        //        {
        //            //Read the first Sheets 
        //            Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
        //            Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
        //            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

        //            foreach (Row row in rows)
        //            {
        //                //Read the first row as header
        //                if (row.RowIndex.Value == 1)
        //                {
        //                    var j = 1;
        //                    foreach (Cell cell in row.Descendants<Cell>())
        //                    {
        //                        var colunmName = firstRowIsHeade ? GetCellValue(doc, cell) : "Field" + j++;
        //                        Console.WriteLine(colunmName);
        //                        Headers.Add(colunmName);
        //                        dt.Columns.Add(colunmName);
        //                    }
        //                }
        //                else
        //                {
        //                    dt.Rows.Add();
        //                    int i = 0;
        //                    foreach (Cell cell in row.Descendants<Cell>())
        //                    {
        //                        dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
        //                        i++;
        //                    }
        //                }
        //            }

        //        }
        //        return dt;
        //    }

            //private string GetCellValue(SpreadsheetDocument doc, Cell cell)
            //{
            //    string value = cell.CellValue.InnerText;
            //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            //    {
            //        return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            //    }
            //    return value;
            //}
        //}
    }
}
