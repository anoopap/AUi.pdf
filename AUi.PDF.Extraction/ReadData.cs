using BitMiracle.Docotic.Pdf;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using itext = iTextSharp.text.pdf;
using Newtonsoft.Json;

namespace AUi.PDF.Extraction
{
    public class ReadData
    {
        private static string pdfFilePath = "";

        public static void DocumentExtract(string file,string template)
        {

            //Delete Existing files
            if (Directory.Exists("Temp"))
            {
                // delete all files in the folder
                string[] files = Directory.GetFiles("Temp");
                foreach (string dfile in files)
                {
                    File.Delete(dfile);
                }

                // delete all subfolders in the folder
                string[] subfolders = Directory.GetDirectories("Temp");
                foreach (string subfolder in subfolders)
                {
                    Directory.Delete(subfolder, true);
                }

                // delete the folder
                Directory.Delete("Temp");
                Console.WriteLine("The folder has been deleted successfully.");
                Directory.CreateDirectory("Temp");
            }
            else
            {
                Directory.CreateDirectory("Temp");
                Console.WriteLine("The folder does not exist.");
            }

            using (itext.PdfReader reader = new itext.PdfReader(file))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    using (iTextSharp.text.Document document = new iTextSharp.text.Document())
                    {
                        using (itext.PdfCopy copy = new itext.PdfCopy(document, new FileStream(string.Format(@"Temp\" + Path.GetFileNameWithoutExtension(file) + @"_{0}.pdf", i), FileMode.Create)))
                        {
                            document.Open();
                            copy.AddPage(copy.GetImportedPage(reader, i));
                            document.Close();
                        }
                    }
                }
            }

            string[] pdfFiles = Directory.GetFiles(@"Temp\", "*.pdf", SearchOption.TopDirectoryOnly);

            var sortedFiles = pdfFiles
                .Select(file => new { File = file, CreationTime = File.GetCreationTime(file) })
                .OrderBy(x => x.CreationTime);
            List<Table> table = new List<Table>();
            List<Texts> text = new List<Texts>();
            foreach (var pdfFile in sortedFiles)
            {

                Console.WriteLine("File: {0}, Created: {1}", pdfFile.File, pdfFile.CreationTime);
                //process the file
                pdfFilePath = pdfFile.File;
                GetCoordinates();
                string Jsonstring = File.ReadAllText(template);
                JObject Jobj = JObject.Parse(Jsonstring);
                IDictionary<string, JToken> Jsondata = JObject.Parse(Jsonstring);

                foreach (KeyValuePair<string, JToken> element in Jsondata)
                {
                    string innerKey = element.Key;
                    JArray a = (JArray)Jobj[element.Key];

                    foreach (var sel in a)
                    {
                        if (sel["Type"].ToString().ToLower() == "text")
                        {
                            var field = GetWholeText(sel["Search Keyword"].ToString(), sel["Position"].ToString(), Convert.ToInt32(sel["Text Width"].ToString()), Convert.ToInt32(sel["Text Height"].ToString()), Convert.ToInt32(sel["Search Text width"].ToString()), Convert.ToInt32(sel["Text Gap"].ToString()), Convert.ToInt32(sel["Move Side"].ToString())).extractedText.Trim();
                            Console.WriteLine(sel["Field Name"].ToString() + " : " + field);
                            Texts txt = new Texts();
                            txt.fieldName = sel["Field Name"].ToString();
                            txt.fieldValue = field;
                            text.Add(txt);
                        }
                        else if (sel["Type"].ToString().ToLower() == "table")
                        {

                            Console.WriteLine(sel["Fields"].ToString());
                            JArray tA = (JArray)sel["Fields"];
                            string tempData = "";
                            foreach (var tableSel in tA)
                            {
                                tempData = tempData + tableSel["Field"].ToString() + "," + tableSel["Field Width"].ToString() + "," + tableSel["Field Adjust"].ToString() + "," + tableSel["Field Splitter"].ToString() + "|";

                            }
                            tempData = tempData.Substring(0, tempData.Length - 1);
                            var outputDt = GetTable(sel["Table Header"].ToString(), sel["Table Footer"].ToString(), Convert.ToInt32(sel["Table Header Width"].ToString()), Convert.ToInt32(sel["Table Footer Width"].ToString()), tempData, sel["Global Replace"].ToString(), Convert.ToInt32(sel["Line Adjust"].ToString()), sel["Remove Rows With Text"].ToString());
                            Table tb = new Table();
                            tb.tableName = sel["Table Name"].ToString();
                            tb.tableValue = outputDt;
                            table.Add(tb);

                        }

                    }
                }




            }
            var tempFields = text.GroupBy(x => x.fieldName)
                                   .Select(x => x.First())
                                   .ToList();
            int fCount = 0;
            foreach (var i in tempFields)
            {

                foreach (var j in text)
                {
                    if (i.fieldName == j.fieldName)
                    {
                        if (String.IsNullOrEmpty(i.fieldValue) && !String.IsNullOrEmpty(j.fieldValue))
                        {
                            tempFields[fCount].fieldValue = j.fieldValue;
                            break;
                        }
                    }

                }
                fCount++;
            }
            DataTable fieldsDt = new DataTable();
            fieldsDt.Columns.Add("Field Name", typeof(string));
            fieldsDt.Columns.Add("Field Value", typeof(string));

            foreach (Texts t in tempFields)
            {
                fieldsDt.Rows.Add(t.fieldName, t.fieldValue);
            }

            List<Table> mergedList = table.GroupBy(x => x.tableName)
                                  .Select(group => new Table
                                  {
                                      tableName = group.Key,
                                      tableValue = group.Aggregate(new DataTable(), (dt, g) => { dt.Merge(g.tableValue); return dt; })
                                  }).ToList();

            // Create a new Excel workbook
            HSSFWorkbook workbook = new HSSFWorkbook();



            /// Add the headers of the DataTable to the header row
            {
                // Create a new sheet
                ISheet sheetTable = workbook.CreateSheet("Data");

                // Create a header row
                IRow headerRowTable = sheetTable.CreateRow(0);
                for (int i = 0; i < fieldsDt.Columns.Count; i++)
                {
                    headerRowTable.CreateCell(i).SetCellValue(fieldsDt.Columns[i].ColumnName);
                }
                ICellStyle cellStyle = workbook.CreateCellStyle();
                cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Aqua.Index;
                cellStyle.FillPattern = FillPattern.SolidForeground;

                for (int i = 0; i < fieldsDt.Columns.Count; i++)
                {
                    ICell cell = headerRowTable.CreateCell(i);
                    cell.SetCellValue(fieldsDt.Columns[i].ColumnName);
                    cell.CellStyle = cellStyle;
                }

                // Add the data of the DataTable to the sheet
                int rowIndex = 1;
                foreach (DataRow row in fieldsDt.Rows)
                {
                    IRow excelRow = sheetTable.CreateRow(rowIndex);

                    for (int i = 0; i < fieldsDt.Columns.Count; i++)
                    {
                        excelRow.CreateCell(i).SetCellValue(row[i].ToString());
                    }

                    rowIndex++;
                }
            }

            //Write Tables
            foreach (var tab in mergedList)
            {
                // Create a new sheet
                ISheet sheetTable = workbook.CreateSheet(tab.tableName);

                // Create a header row
                IRow headerRowTable = sheetTable.CreateRow(0);


                // Add the headers of the DataTable to the header row
                for (int i = 0; i < tab.tableValue.Columns.Count; i++)
                {
                    headerRowTable.CreateCell(i).SetCellValue(tab.tableValue.Columns[i].ColumnName);
                }
                ICellStyle cellStyle = workbook.CreateCellStyle();
                cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Aqua.Index;
                cellStyle.FillPattern = FillPattern.SolidForeground;

                for (int i = 0; i < tab.tableValue.Columns.Count; i++)
                {
                    ICell cell = headerRowTable.CreateCell(i);
                    cell.SetCellValue(tab.tableValue.Columns[i].ColumnName);
                    cell.CellStyle = cellStyle;
                }

                // Add the data of the DataTable to the sheet
                int rowIndex = 1;
                foreach (DataRow row in tab.tableValue.Rows)
                {
                    IRow excelRow = sheetTable.CreateRow(rowIndex);

                    for (int i = 0; i < tab.tableValue.Columns.Count; i++)
                    {
                        excelRow.CreateCell(i).SetCellValue(row[i].ToString());
                    }

                    rowIndex++;
                }

            }
            // Write the workbook to a file
            using (FileStream stream = new FileStream(Path.GetDirectoryName(file) + @"\" + Path.GetFileNameWithoutExtension(file) + ".xls", FileMode.Create))
            {
                workbook.Write(stream);
            }

            //Writer a json file


            var result = "{";
            for (int i = 0; i < fieldsDt.Rows.Count; i++)
            {
                result += "\"" + fieldsDt.Rows[i][0].ToString() + "\":\"" + fieldsDt.Rows[i][1].ToString().Replace(Environment.NewLine, " ").Replace(@"\r\n", " ").Replace("\"", "\\\"") + "\",";
            }

            foreach (var tables in mergedList)
            {
                result += "\"" + tables.tableName + "\":[";
                foreach (DataRow row in tables.tableValue.Rows)
                {
                    result += "{";
                    for (int i = 0; i < tables.tableValue.Columns.Count; i++)
                    {
                        result += "\"" + tables.tableValue.Columns[i].ColumnName + "\":\"" + row[i].ToString().Replace(Environment.NewLine, " ").Replace(@"\r\n", " ").Replace("\"", "\\\"") + "\",";
                    }
                    result = result.TrimEnd(',');
                    result += "},";
                }
                result = result.TrimEnd(',');
                result += "],";
            }
            result = result.TrimEnd(',');
            result += "}";
            File.WriteAllText(Path.GetDirectoryName(file) + @"\" + Path.GetFileNameWithoutExtension(file) + ".json", result);

            File.WriteAllText(Path.GetDirectoryName(file) + @"\" + Path.GetFileNameWithoutExtension(file) + ".json", JsonConvert.SerializeObject(JsonConvert.DeserializeObject(result), Formatting.Indented));

        }
        private static DataTable GetTable(string tableHeader, string tableFooters, int tableHeaderWidth, int tableFooterWidth, string headers, string globalReplace = "", int lineAdjust = 0, string removeRowsWithText = "")
        {
            double headerY = 0;
            double footerY = 0;
            double headerIgnore = 0;
            DataTable outputTable = new DataTable();
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                PdfPage page = pdf.Pages[0];

                //Find Header Coordinates
                foreach (PdfTextData data in page.GetWords())
                {
                    if (tableHeader.Split(' ')[0] == data.GetText())
                    {
                        Console.WriteLine(data.Bounds.Height);
                        if (GetTextByCoordinates(data.Bounds, tableHeaderWidth).ToLower().Trim().Contains(tableHeader.ToLower().Trim()))
                        {
                            headerY = data.Bounds.Y;
                            headerIgnore = data.Bounds.Height;
                            Console.WriteLine("Header Y : " + headerY);
                            break;
                        }

                    }

                }
                //Find Footer Coordinates
                foreach (string tableFooter in tableFooters.Split('|'))
                {
                    foreach (PdfTextData data in page.GetWords())
                    {
                        if (tableFooter.Split(' ')[0] == data.GetText())
                        {
                            if (GetTextByCoordinates(data.Bounds, tableFooterWidth).ToLower().Trim().Contains(tableFooter.ToLower().Trim()))
                            {
                                footerY = data.Bounds.Y;
                                Console.WriteLine("Footer Y : " + footerY);
                                break;
                            }

                        }

                    }
                }
                if (footerY == 0)
                {
                    footerY = page.Height;
                }
                // Console.WriteLine(GetTextByCoordinates(new PdfRectangle(0,headerY+ headerIgnore, page.Width,(footerY-headerY- headerIgnore)),0));
                var tableRectangle = new PdfRectangle(0, headerY + headerIgnore, page.Width, (footerY - headerY - headerIgnore));


                List<string> holdRows = new List<string>();
                foreach (string header in headers.Split('|'))
                {
                    outputTable.Columns.Add(header.Split(',')[0].Trim(), typeof(string));
                }
                string masterColumnName = "";

                List<extractionData> edBreak = new List<extractionData>();
                foreach (string header in headers.Split('|'))
                {
                    if (header.ToLower().Contains("true"))
                    {
                        foreach (PdfTextData data in page.GetWords())
                        {
                            if (header.Split(',')[0].Split(' ')[0] == data.GetText())
                            {
                                if (GetTextByCoordinates(new PdfRectangle(data.Bounds.X, data.Bounds.Y, (page.Width - data.Bounds.X), data.Bounds.Height), 0).ToLower().Contains(header.Split(',')[0].ToLower()))
                                {

                                    var tempHeader = GetWholeText(header.Split(',')[0], "bottom", Convert.ToDouble(header.Split(',')[1]), (footerY - headerY - headerIgnore - 10 - data.Bounds.Height), Convert.ToDouble(header.Split(',')[1]), move: Convert.ToDouble(header.Split(',')[2]));

                                    Console.WriteLine(tempHeader.coordinates.X.ToString() + " " + (tempHeader.coordinates.Y - 10).ToString() + " " + tempHeader.coordinates.Width + " " + tempHeader.coordinates.Height);


                                    edBreak = GetTextListByCoordinates(new PdfRectangle(tempHeader.coordinates.X, tempHeader.coordinates.Y - 10, tempHeader.coordinates.Width, tempHeader.coordinates.Height));
                                    masterColumnName = header.Split(',')[0].Trim();
                                    Console.WriteLine(tempHeader.extractedText);
                                    break;
                                }

                            }
                        }
                        break;
                    }
                }

                double[] array = edBreak.Select(x => x.coordinates.Y).ToArray();
                Array.Sort(array);
                double tempDiff = 0;
                for (int i = 0; i <= array.Length - 1; i++)
                {
                    DataRow itemRow = outputTable.NewRow();
                    double difference = 0;
                    if (array.Length == 1)
                    {
                        difference = 0;
                    }
                    else
                    {
                        try
                        {
                            difference = array[i + 1] - array[i];
                        }
                        catch (Exception e)
                        {
                            difference = footerY - array[i];
                        }

                    }

                    // Console.WriteLine(difference);
                    if (difference <= 0)
                    {
                        difference = footerY - array[i];
                    }
                    foreach (string header in headers.Split('|'))
                    {
                        double lineX = 0;
                        double lineY = 0;
                        double lineWidth = 0;

                        foreach (PdfTextData data in page.GetWords())
                        {
                            if (header.Split(',')[0].Split(' ')[0] == data.GetText())
                            {
                                if (GetTextByCoordinates(new PdfRectangle(data.Bounds.X, data.Bounds.Y, (page.Width - data.Bounds.X), data.Bounds.Height), 0).ToLower().Contains(header.Split(',')[0].ToLower()))
                                {
                                    var tempHeader = GetWholeText(header.Split(',')[0], "bottom", Convert.ToDouble(header.Split(',')[1]), (footerY - headerY - headerIgnore - 10 - data.Bounds.Height), Convert.ToDouble(header.Split(',')[1]), move: Convert.ToDouble(header.Split(',')[2]));
                                    lineX = tempHeader.coordinates.X;
                                    lineY = tempHeader.coordinates.Y + tempDiff + lineAdjust;
                                    lineWidth = tempHeader.coordinates.Width;

                                    Console.WriteLine(GetTextByCoordinates(new PdfRectangle(lineX, lineY, lineWidth, difference - 5), lineWidth));
                                    itemRow[header.Split(',')[0].Trim()] = GetTextByCoordinates(new PdfRectangle(lineX, lineY, lineWidth, difference - 5), lineWidth);

                                    break;
                                }

                            }
                        }

                    }
                    tempDiff = tempDiff + difference;
                    outputTable.Rows.Add(itemRow);
                }

                //Process Datatable
                try
                {
                    foreach (DataRow row in outputTable.Rows)
                    {
                        foreach (DataColumn column in outputTable.Columns)
                        {
                            if (row[column] != DBNull.Value)
                            {
                                row[column] = row[column].ToString().Replace(globalReplace, "").Trim();
                            }
                        }
                    }
                }
                catch (Exception e)
                {

                }


                outputTable.AsEnumerable()
    .Where(r => r.ItemArray.All(f => f is DBNull || string.IsNullOrWhiteSpace(f.ToString())))
    .ToList()
    .ForEach(r => outputTable.Rows.Remove(r));



                for (int i = 1; i < outputTable.Rows.Count; i++)
                {
                    bool isEmpty = false;
                    if (string.IsNullOrEmpty(outputTable.Rows[i][masterColumnName].ToString()))
                    {
                        isEmpty = true;
                    }
                    if (isEmpty)
                    {
                        for (int j = 0; j < outputTable.Columns.Count; j++)
                        {
                            outputTable.Rows[i - 1][j] += " " + outputTable.Rows[i][j].ToString();
                        }
                        outputTable.Rows[i].Delete();
                    }
                }
                outputTable.AcceptChanges();

                if (!String.IsNullOrEmpty(removeRowsWithText))
                {
                    outputTable.AsEnumerable()
    .Where(row => row.ItemArray.Any(col => col.ToString().Contains(removeRowsWithText)))
    .ToList()
    .ForEach(row => row.Delete());
                    outputTable.AcceptChanges();
                }


                Console.WriteLine(GetTextByCoordinates(new PdfRectangle(0, headerY + headerIgnore, page.Width, (footerY - headerY - headerIgnore)), 0));

            }
            return outputTable;
        }
        private static extractionData GetWholeText(string searchText, string position, double width, double height, double searchTextWidth, double gap = 10, double move = 0)
        {
            string returnText = "";
            extractionData ed = new extractionData();
            ed.extractedText = "";
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                PdfPage page = pdf.Pages[0];



                foreach (PdfTextData data in page.GetWords())
                {
                    if (searchText.Split(' ')[0] == data.GetText())
                    {
                        if (GetTextByCoordinates(data.Bounds, searchTextWidth).ToLower().Trim().Contains(searchText.ToLower().Trim()))
                        {
                            // Console.WriteLine(GetTextByCoordinates(data.Bounds, searchTextWidth));

                            switch (position)
                            {
                                case "left":
                                    returnText = GetTextByCoordinates(new PdfRectangle((float)(data.Bounds.X - width - gap + move), (float)(data.Bounds.Y), (float)width, (float)height), width);
                                    ed.coordinates = new PdfRectangle((float)(data.Bounds.X - width - gap + move), (float)(data.Bounds.Y), (float)width, (float)height);
                                    ed.extractedText = returnText.Replace(searchText.Trim(), "");
                                    break;
                                case "right":
                                    returnText = GetTextByCoordinates(new PdfRectangle((float)(data.Bounds.X + data.Bounds.Width + gap + move), (float)(data.Bounds.Y), (float)width, (float)height), width);
                                    ed.coordinates = new PdfRectangle((float)(data.Bounds.X + searchTextWidth + gap + move), (float)(data.Bounds.Y), (float)width, (float)height);
                                    ed.extractedText = returnText.Replace(searchText.Trim(), "");
                                    break;
                                case "top":
                                    returnText = GetTextByCoordinates(new PdfRectangle((float)(data.Bounds.X - (width / 2) + move), (float)(data.Bounds.Y - gap - height), (float)width, (float)height), width);
                                    ed.coordinates = new PdfRectangle((float)(data.Bounds.X - (width / 2) + move), (float)(data.Bounds.Y - gap - height), (float)width, (float)height);
                                    ed.extractedText = returnText.Replace(searchText.Trim(), "");
                                    break;
                                case "bottom":
                                    returnText = GetTextByCoordinates(new PdfRectangle((float)(data.Bounds.X - (width / 2) + move), (float)(data.Bounds.Y + data.Bounds.Height + gap), (float)width, (float)height), width);
                                    ed.coordinates = new PdfRectangle((float)(data.Bounds.X - (width / 2) + move), (float)(data.Bounds.Y + data.Bounds.Height + gap), (float)width, (float)height);
                                    ed.extractedText = returnText.Replace(searchText.Trim(), "");
                                    break;
                                default:
                                    // code block
                                    break;
                            }
                            break;
                        }

                    }

                }
            }


            return ed;
        }
        private static string GetTextByCoordinates(PdfRectangle rectangle, double searchTextWidth)
        {
            string areaText = "";
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                var page = pdf.Pages[0];

                if (searchTextWidth == 0)
                {
                    searchTextWidth = page.Width - rectangle.X;
                }
                var options = new PdfTextExtractionOptions
                {
                    Rectangle = new PdfRectangle(rectangle.X, rectangle.Y, searchTextWidth, rectangle.Height),
                    WithFormatting = false

                };
                areaText = page.GetText(options);
            }
            return areaText;
        }
        private static List<extractionData> GetTextListByCoordinates(PdfRectangle rectangle)
        {
            List<extractionData> edList = new List<extractionData>();
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                var page = pdf.Pages[0];

                foreach (PdfTextData data in page.GetWords())
                {
                    if (((data.Bounds.X >= rectangle.X) && (data.Bounds.X <= rectangle.X + rectangle.Width)) && ((data.Bounds.Y >= rectangle.Y) && (data.Bounds.Y <= (rectangle.Y + rectangle.Height))))
                    {                       
                        extractionData ed = new extractionData();
                        ed.coordinates = data.Bounds;
                        ed.extractedText = data.GetText();
                        edList.Add(ed);
                    }
                }
            }
            return edList;
        }
        private static void GetCoordinates()
        {
            List<extractionData> edList = new List<extractionData>();
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                var page = pdf.Pages[0];

                foreach (PdfTextData data in page.GetWords())
                {
                    Console.WriteLine(
        $"{{\n" +
        $"  text: '{data.GetText()}',\n" +
        $"  bounds: {data.Bounds},\n" +
        $"  font name: '{data.Font.Name}',\n" +
        $"  font size: {data.FontSize},\n" +
        $"  transformation matrix: {data.TransformationMatrix},\n" +
        $"  rendering mode: '{data.RenderingMode}',\n" +
        $"  brush: {data.Brush},\n" +
        $"  pen: {data.Pen}\n" +
        $"}},"
         );
                }
            }
        }
    }
    
    public class extractionData
    {
        public string extractedText { get; set; }
        public PdfRectangle coordinates { get; set; }
    }
    public class Texts
    {
        public string fieldName { get; set; }
        public string fieldValue { get; set; }
    }
    public class Table
    {
        public string tableName { get; set; }
        public DataTable tableValue { get; set; }
    }
}
