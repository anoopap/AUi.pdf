using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using BitMiracle.Docotic.Pdf;
using itext = iTextSharp.text.pdf;
using PdfDocument = BitMiracle.Docotic.Pdf.PdfDocument;

namespace AUi
{
    public class PDF
    {
        public static void MergePDF(string FolderPath, string outputPdfPath)
        {
            string[] lstFiles = Directory.GetFiles(FolderPath);
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage;


            sourceDocument = new Document();
            pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

            //Open the output file
            sourceDocument.Open();

            try
            {
                //Loop through the files list
                try
                {
                    for (int f = 0; f < lstFiles.Length; f++)
                    {
                        int pages = GetPDFCount(lstFiles[f]);

                        reader = new PdfReader(lstFiles[f]);
                        //Add pages of current file
                        for (int i = 1; i <= pages; i++)
                        {
                            importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                            pdfCopyProvider.AddPage(importedPage);
                        }

                        reader.Close();
                    }
                }
                catch (Exception)
                {

                }
                //At the end save the output file
                sourceDocument.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public static int GetPDFCount(string file)
        {
            using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(file)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }
    }
    public class ReadData
    {
        private static string pdfFilePath = "";

        public static ExtractedData DocumentExtract(string file, string template, bool generatefile = false)
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
                            var field = GetWholeText(sel["Search Keyword"].ToString(), sel["Position"].ToString(), Convert.ToInt32(sel["Text Width"].ToString()), Convert.ToInt32(sel["Text Height"].ToString()), Convert.ToInt32(sel["Search Text width"].ToString()), Convert.ToInt32(sel["Move Vertical"].ToString()), Convert.ToInt32(sel["Text Gap"].ToString()), Convert.ToInt32(sel["Move Side"].ToString()), Convert.ToInt32(sel["Search Text Height"].ToString()), Convert.ToInt32(sel["Search Text AdjustVertical"].ToString()), Convert.ToInt32(sel["Search Text AdjustSide"].ToString())).extractedText.Trim();
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
                                tempData = tempData + tableSel["Field"].ToString() + "," + tableSel["Field Width"].ToString() + "," + tableSel["Field Adjust"].ToString() + "," + tableSel["Field Splitter"].ToString() + "," + tableSel["SkipWord"].ToString() + "|";
                            }
                            tempData = tempData.Substring(0, tempData.Length - 1);
                            var outputDt = GetTable(sel["Table Header"].ToString(), sel["Table Footer"].ToString(), Convert.ToInt32(sel["Table Header Width"].ToString()), Convert.ToInt32(sel["Table Footer Width"].ToString()), tempData, sel["Global Replace"].ToString(), Convert.ToInt32(sel["Line Adjust"].ToString()), sel["Remove Rows With Text"].ToString(), Convert.ToInt32(sel["Skip Header Word"].ToString()), Convert.ToInt32(sel["Skip Footer Word"].ToString()));
                            Table tb = new Table();
                            tb.tableName = sel["Table Name"].ToString();
                            tb.tableValue = outputDt;
                            table.Add(tb);

                        }
                        else if (sel["Type"].ToString().ToLower() == "singleheadertable")
                        {
                            Console.WriteLine(sel["Fields"].ToString());
                            JArray tA = (JArray)sel["Fields"];
                            string tempData = "";
                            foreach (var tableSel in tA)
                            {
                                tempData = tempData + tableSel["Field"].ToString() + "," + tableSel["Field Width"].ToString() + "," + tableSel["Field Adjust"].ToString() + "," + tableSel["Field Splitter"].ToString() + "," + tableSel["SkipWord"].ToString() + "|";
                            }
                            tempData = tempData.Substring(0, tempData.Length - 1);
                            var outputDt = SingleHeaderTable(sel["Table Header"].ToString(), sel["Table Footer"].ToString(), Convert.ToInt32(sel["Table Header Width"].ToString()), Convert.ToInt32(sel["Table Footer Width"].ToString()), tempData, Convert.ToInt32(sel["CellWidth"].ToString()), sel["Global Replace"].ToString(), Convert.ToInt32(sel["Line Adjust"].ToString()), sel["Remove Rows With Text"].ToString(), Convert.ToInt32(sel["Skip Header Word"].ToString()), Convert.ToInt32(sel["Skip Footer Word"].ToString()));
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

            if (generatefile)
            {

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
            List<DataTable> new_table = new List<DataTable>();
            foreach (Table in_table in mergedList)
            {
                DataTable temptable = new DataTable();
                temptable.TableName = in_table.tableName;
                foreach (DataColumn column in in_table.tableValue.Columns)
                {
                    temptable.Columns.Add(column.ColumnName);
                }
                foreach (DataRow row in in_table.tableValue.Rows)
                {
                    temptable.Rows.Add(row.ItemArray);
                }
                new_table.Add(temptable);
            }
            ExtractedData extractedata = new ExtractedData();
            extractedata.textfields = fieldsDt;
            extractedata.tables = new_table;
            return (extractedata);
        }
        private static DataTable GetTable(string tableHeader, string tableFooters, int tableHeaderWidth, int tableFooterWidth, string headers, string globalReplace = "", int lineAdjust = 0, string removeRowsWithText = "", int skipHeaderWord = 0, int SkipFooterWord = 0)
        {
            double headerY = 0;
            double footerY = 0;
            double headerIgnore = 0;
            bool headerfound = false;
            DataTable outputTable = new DataTable();
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                BitMiracle.Docotic.Pdf.PdfPage page = pdf.Pages[0];

                //Find Header Coordinates
                foreach (PdfTextData data in page.GetWords())
                {
                    if (tableHeader.Split(' ')[0] == data.GetText())
                    {
                        if (skipHeaderWord > 0)
                        {
                            skipHeaderWord--;
                            continue;
                        }
                        Console.WriteLine(data.Bounds.Height);
                        if (GetTextByCoordinates(BitMiracle.Docotic.Pdf.PdfTextData data.Bounds, tableHeaderWidth).ToLower().Trim().Replace(" ", "").Contains(tableHeader.ToLower().Trim().Replace(" ", "")))
                        {
                            headerY = data.Bounds.Y;
                            headerIgnore = data.Bounds.Height;
                            Console.WriteLine("Header Y : " + headerY);
                            headerfound = true;
                            break;
                        }

                    }

                }
                if (!headerfound)
                {
                    return (new DataTable());
                }
                //Find Footer Coordinates
                foreach (string tableFooter in tableFooters.Split('|'))
                {
                    foreach (PdfTextData data in page.GetWords())
                    {
                        if (tableFooter.Split(' ')[0] == data.GetText())
                        {
                            if (SkipFooterWord > 0)
                            {
                                SkipFooterWord--;
                                continue;
                            }
                            if (GetTextByCoordinates(data.Bounds, tableFooterWidth).ToLower().Trim().Replace(" ", "").Contains(tableFooter.ToLower().Trim().Replace(" ", "")))
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
                    int skipcount, skipcount2 = 0;
                    skipcount = Convert.ToInt32(header.Split(',')[4].ToString());
                    skipcount2 = skipcount;
                    if (header.ToLower().Contains("true"))
                    {
                        foreach (PdfTextData data in page.GetWords())
                        {
                            if (header.Split(',')[0].Split(' ')[0] == data.GetText())
                            {
                                if (skipcount > 0)
                                {
                                    skipcount--;
                                    continue;
                                }
                                if (GetTextByCoordinates(new PdfRectangle(data.Bounds.X, data.Bounds.Y, (page.Width - data.Bounds.X), data.Bounds.Height), 0).ToLower().Replace(" ", "").Contains(header.Split(',')[0].ToLower().Replace(" ", "")))
                                {

                                    var tempHeader = GetWholeText(header.Split(',')[0], "bottom", Convert.ToDouble(header.Split(',')[1]), (footerY - headerY - headerIgnore - data.Bounds.Height), Convert.ToDouble(header.Split(',')[1]), move: Convert.ToDouble(header.Split(',')[2]), skipcount: skipcount2);

                                    Console.WriteLine(tempHeader.coordinates.X.ToString() + " " + (tempHeader.coordinates.Y - 10).ToString() + " " + tempHeader.coordinates.Width + " " + tempHeader.coordinates.Height);


                                    edBreak = GetTextListByCoordinates(new PdfRectangle(tempHeader.coordinates.X - 2, tempHeader.coordinates.Y - 10, tempHeader.coordinates.Width, tempHeader.coordinates.Height));
                                    masterColumnName = header.Split(',')[0].Trim();
                                    Console.WriteLine(tempHeader.extractedText);
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }

                //To find the lines
                double[] array = edBreak.Select(x => x.coordinates.Y).ToArray();
                Array.Sort(array);
                double tempDiff = 0;
                double prevliney = 0;

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

                    int columncounter = 0;

                    //where extraction takes place
                    foreach (string header in headers.Split('|'))
                    {
                        columncounter++;
                        int skipcount, skipcount2 = 0;
                        skipcount = Convert.ToInt32(header.Split(',')[4].ToString());
                        skipcount2 = skipcount;
                        double lineX = 0;
                        double lineY = 0;
                        double lineWidth = 0;

                        foreach (PdfTextData data in page.GetWords())
                        {
                            if (header.Split(',')[0].Split(' ')[0] == data.GetText())
                            {
                                if (skipcount > 0)
                                {
                                    skipcount--;
                                    continue;
                                }
                                if ((columncounter > 1 && data.Bounds.Y + 30 < prevliney) && footerY != page.Height)
                                {
                                    continue;
                                }
                                if (GetTextByCoordinates(new PdfRectangle(data.Bounds.X, data.Bounds.Y, (page.Width - data.Bounds.X), data.Bounds.Height), 0).Replace(" ", "").ToLower().Contains(header.Split(',')[0].ToLower().Replace(" ", "")))
                                {
                                    var tempHeader = GetWholeText(header.Split(',')[0], "bottom", Convert.ToDouble(header.Split(',')[1]), (footerY - headerY - headerIgnore - 10 - data.Bounds.Height), Convert.ToDouble(header.Split(',')[1]), move: Convert.ToDouble(header.Split(',')[2]), skipcount: skipcount2);
                                    lineX = tempHeader.coordinates.X;
                                    lineY = tempHeader.coordinates.Y + tempDiff + lineAdjust;
                                    lineWidth = tempHeader.coordinates.Width;
                                    prevliney = lineY;

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
        private static extractionData GetWholeText(string searchText, string position, double width, double height, double searchTextWidth, double movevertical = 0, double gap = 10, double move = 0, double searchTextHeight = 0, double adjustSearchVertical = 0, double adjustSearchHorizontal = 0, int skipcount = 0)
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
                        if (skipcount > 0)
                        {
                            skipcount--;
                            continue;
                        }
                        if (GetTextByCoordinates(data.Bounds, searchTextWidth, searchTextHeight, adjustSearchVertical, adjustSearchHorizontal).ToLower().Trim().Replace(" ", "").Contains(searchText.ToLower().Trim().Replace(" ", "")))
                        {
                            // Console.WriteLine(GetTextByCoordinates(data.Bounds, searchTextWidth));

                            switch (position)
                            {
                                case "left":
                                    returnText = GetTextByCoordinates(new PdfRectangle((float)(data.Bounds.X - width - gap + move), (float)(data.Bounds.Y + movevertical), (float)width, (float)height), width);
                                    ed.coordinates = new PdfRectangle((float)(data.Bounds.X - width - gap + move), (float)(data.Bounds.Y), (float)width, (float)height);
                                    ed.extractedText = returnText.Replace(searchText.Trim(), "");
                                    break;
                                case "right":
                                    returnText = GetTextByCoordinates(new PdfRectangle((float)(data.Bounds.X + data.Bounds.Width + gap + move), (float)(data.Bounds.Y + movevertical), (float)width, (float)height), width);
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
        private static string GetTextByCoordinates(PdfRectangle rectangle, double searchTextWidth, double searchTextHeight = 0, double adjustSearchVertical = 0, double adjustSearchHorizontal = 0)
        {
            string areaText = "";
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                var page = pdf.Pages[0];

                if (searchTextWidth == 0)
                {
                    searchTextWidth = page.Width - rectangle.X;
                }
                searchTextHeight = rectangle.Height + searchTextHeight;

                var options = new PdfTextExtractionOptions
                {
                    Rectangle = new PdfRectangle(rectangle.X + adjustSearchHorizontal, rectangle.Y + adjustSearchVertical, searchTextWidth, searchTextHeight),
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

        public static bool TextExists(string pdfFilePath, string searchText)
        {
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                foreach (PdfPage page in pdf.Pages)
                {
                    foreach (PdfTextData data in page.GetWords())
                    {
                        if (searchText.Split(' ')[0] == data.GetText())
                        {
                            return (true);
                        }
                    }
                }
            }
            return (false);
        }

        private static DataTable SingleHeaderTable(string tableHeader, string tableFooters, int tableHeaderWidth, int tableFooterWidth, string headers, int cellwidth, string globalReplace = "", int lineAdjust = 0, string removeRowsWithText = "", int skipHeaderWord = 0, int SkipFooterWord = 0)
        {
            double headerY = 0;
            double footerY = 0;
            double headerIgnore = 0;
            bool headerfound = false;
            DataTable outputTable = new DataTable();
            using (var pdf = new PdfDocument(pdfFilePath))
            {
                PdfPage page = pdf.Pages[0];

                //Find Header Coordinates
                foreach (PdfTextData data in page.GetWords())
                {
                    if (tableHeader.Split(' ')[0] == data.GetText())
                    {
                        if (skipHeaderWord > 0)
                        {
                            skipHeaderWord--;
                            continue;
                        }
                        Console.WriteLine(data.Bounds.Height);
                        if (GetTextByCoordinates(data.Bounds, tableHeaderWidth).ToLower().Trim().Contains(tableHeader.ToLower().Trim()))
                        {
                            headerY = data.Bounds.Y;
                            headerIgnore = data.Bounds.Height;
                            Console.WriteLine("Header Y : " + headerY);
                            headerfound = true;
                            break;
                        }

                    }

                }
                if (!headerfound)
                {
                    return (new DataTable());
                }
                //Find Footer Coordinates
                foreach (string tableFooter in tableFooters.Split('|'))
                {
                    foreach (PdfTextData data in page.GetWords())
                    {
                        if (tableFooter.Split(' ')[0] == data.GetText())
                        {
                            if (SkipFooterWord > 0)
                            {
                                SkipFooterWord--;
                                continue;
                            }
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
                    int skipcount, skipcount2 = 0;
                    skipcount = Convert.ToInt32(header.Split(',')[4].ToString());
                    skipcount2 = skipcount;
                    if (header.ToLower().Contains("true"))
                    {
                        foreach (PdfTextData data in page.GetWords())
                        {
                            if (header.Split(',')[0].Split(' ')[0] == data.GetText())
                            {
                                if (skipcount > 0)
                                {
                                    skipcount--;
                                    continue;
                                }
                                if (GetTextByCoordinates(new PdfRectangle(data.Bounds.X, data.Bounds.Y, (page.Width - data.Bounds.X), data.Bounds.Height), 0).ToLower().Replace(" ", "").Contains(header.Split(',')[0].ToLower().Replace(" ", "")))
                                {

                                    var tempHeader = GetWholeText(header.Split(',')[0], "bottom", Convert.ToDouble(header.Split(',')[1]), (footerY - headerY - headerIgnore - 10 - data.Bounds.Height), Convert.ToDouble(header.Split(',')[1]), move: Convert.ToDouble(header.Split(',')[2]), skipcount: skipcount2);

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

                //To find the lines
                double[] array = edBreak.Select(x => x.coordinates.Y).ToArray();
                Array.Sort(array);
                double tempDiff = 0;
                double prevliney = 0;

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
                        difference = array[1] - array[0];
                    }
                    catch (Exception e)
                    {
                        try
                        {
                            difference = footerY - array[0];
                        }
                        catch
                        {
                            difference = 10;
                        }
                    }


                }

                // Console.WriteLine(difference);
                if (difference <= 0)
                {
                    difference = footerY - array[0];
                }

                foreach (extractionData ed in edBreak)
                {
                    if (ed.coordinates.Y > footerY)
                    {
                        break;
                    }
                    int columncounter = 0;
                    List<string> row = new List<string>();
                    foreach (string header in headers.Split('|'))
                    {
                        row.Add(GetTextByCoordinates(new PdfRectangle(ed.coordinates.X + (columncounter * cellwidth) + Convert.ToDouble(header.Split(',')[2]), ed.coordinates.Y, Convert.ToDouble(header.Split(',')[1]), ed.coordinates.Height), Convert.ToDouble(header.Split(',')[1])));
                        columncounter++;
                    }
                    outputTable.Rows.Add(row.ToArray());
                }

                tempDiff = tempDiff + difference;
                outputTable.Rows.Add(itemRow);

            }
            return outputTable;
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
    public class ExtractedData
    {
        public DataTable textfields { get; set; }
        public List<DataTable> tables { get; set; }
    }
}
