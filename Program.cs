//TODO: ML caricare più file da singola directory.
//MLContext mlcontext = new MLContext
//IDataView data = mlcontext.Data.LoadFromTextFile<HousingData>(@"C:\Users\Giorgio Della Roscia\Desktop\ML\Progetti\SautinSoft\PdfToCsv\TXT\*", separatorChar: ',', hasHeader: true);
using iTextSharp.text;
using iTextSharp.text.pdf;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PdfToCsv
{
    class Program
    {
        public const string projectdirpath = @"C:\Users\Giorgio Della Roscia\Desktop\ML\Progetti\SautinSoft\PdfToCsv\";
        static void Main(string[] args)
        {
            List<FileInfo> pdffiles = GetPdfFilePathList("PDF");
            //var xlsfileindex = 0;
            CreateTxtFilesHavingXls(); //usare se ho già creato i miei xls file
            //foreach (var pdffile in pdffiles)
            //{
            //    try
            //    {
            //        List<string> listpdfsplitted = SplitPdfFileInSinglePage(pdffile);
            //        CreateTxtFilesWithoutHavingXls(listpdfsplitted); //usare se devo creare xlsfile
            //    }
            //    catch (ArgumentException)
            //    {
            //        Console.WriteLine($"AE: {pdffile.Name}.");
            //        //File.Delete(pdffile.FullName);
            //        //File.Delete(xlspathlist[xlsfileindex]);
            //    }
            //    catch (Exception)
            //    {
            //        Console.WriteLine($"E: {pdffile.Name}.");
            //        //File.Delete(pdffile.FullName);
            //        //File.Delete(xlspathlist[xlsfileindex]);
            //    }
            //    //xlsfileindex++;
            //}
            Console.ReadLine();
        }

        private static List<FileInfo> GetPdfFilePathList(string folder)
        {
            DirectoryInfo pdfpathlist = new DirectoryInfo($@"{projectdirpath}\{folder}");
            return pdfpathlist.GetFiles().ToList();
        }

        private static List<string> SplitPdfFileInSinglePage(FileInfo file)
        {
            List<string> filelist = new List<string>();
            string newfullname = "";
            using (PdfReader pdfreader = new PdfReader(file.FullName))
            {
                for (int pagenumber = 0; pagenumber < pdfreader.NumberOfPages; pagenumber++)
                {
                    string newname = file.Name.Replace(".pdf", "");
                    newfullname = string.Format($@"{projectdirpath}SplittedPDF\{newname}_page{pagenumber}");

                    Document document = new Document();
                    PdfCopy copy = new PdfCopy(document, new FileStream($"{newfullname}.pdf", FileMode.Create));
                    document.Open();

                    if (pagenumber < pdfreader.NumberOfPages)
                    {
                        copy.AddPage(copy.GetImportedPage(pdfreader, pagenumber + 1));
                    }
                    document.Close();
                    filelist.Add(newfullname);
                }
            }
            return filelist;
        }

        private static List<string> GetXlsFilePathList()
        {
            List<string> xlsfilesString = new List<string>();
            List<FileInfo> xlsfilesFileInfo = GetPdfFilePathList("XLS");
            foreach (var xlsfile in xlsfilesFileInfo) //lascio il foreach per utilizzare una lista di stringhe anzichè di FileInfo
            {
                try
                {
                    xlsfilesString.Add($"{xlsfile}");
                }
                catch
                {
                    Console.WriteLine($"Return path list ERROR: {xlsfile.Name}.");
                }
            }
            return xlsfilesString;
        }

        private static List<string> CreateXlsFile(List<string> listPdfFiles)
        {
            List<string> fileList = new List<string>();
            listPdfFiles.ForEach(pdfFileName =>
            {
                SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
                string xlsfullfame = pdfFileName.Replace(@"\SplittedPDF\", @"\XLS\"); //li copio in xlsdir
                try
                {
                    f.OpenPdf($"{pdfFileName}.pdf"); //apro i "vecchi" pdf
                    f.ToExcel($"{xlsfullfame}.xls"); //trasformo i nuovi file senza estensione in xls
                    string compactpathxls = $"{xlsfullfame}.xls";
                    fileList.Add(compactpathxls);
                }
                catch
                {
                    Console.WriteLine($"Read file ERROR: {pdfFileName}.");
                }
            });
            return fileList;
        }

        private static void CreateTxtFilesWithoutHavingXls(List<string> listpdfsplitted)
        {
            List<string> xlsfiles = CreateXlsFile(listpdfsplitted);
            foreach (string xlsfile in xlsfiles)
            {
                string txtfilename = ExtrapolateFileName(xlsfile);
                if (txtfilename != null && txtfilename.Contains('-'))
                {
                    CreateTxtFile(txtfilename, xlsfile);
                }
            }
        }

        private static void CreateTxtFilesHavingXls()
        {
            List<string> xlsfilelist = GetXlsFilePathList();
            foreach (string xlsfile in xlsfilelist)
            {
                string txtfilename = ExtrapolateFileName(xlsfile);
                if (txtfilename != null && txtfilename.Contains('-'))
                {
                    CreateTxtFile(txtfilename, xlsfile);
                }
            }
        }

        private static string ExtrapolateFileName(string fullxlspath)
        {
            HSSFWorkbook hssfworkbook;
            using (FileStream excelfile = new FileStream(fullxlspath, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(excelfile);
            }
            ISheet sheet = hssfworkbook.GetSheetAt(0);
            string producerdata = sheet.GetRow(6).GetCell(0).ToString();
            string producerIDname = GetProducerIDName(producerdata);
            return producerIDname;
        }

        private static string GetProducerIDName(string producerdata)
        {
            string producerIDname = "";
            if (producerdata != "")
            {
                //TODO: rimuovere i file con formato errato
                RegexOptions options = RegexOptions.None;
                Regex regex = new Regex("[ ]{2,}", options);
                int startindex = producerdata.IndexOf(':') + 2;
                int endindex = producerdata.IndexOf('a') - 2; //non metto \n altrimenti prende quello dopo "Produttore:" e neanche 'L' perchè può esserci nel nome
                producerIDname = producerdata.Substring(startindex, endindex - startindex);
                producerIDname = regex.Replace(producerIDname.Replace("\n", " "), " ");
                producerIDname = producerIDname.Replace(".", "");
            }
            return producerIDname;
        }

        private static void CreateTxtFile(string txtfilename, string xlsfilepath)
        {
            string txtfilepath = $@"{projectdirpath}TXT\{txtfilename}.txt";
            using (StreamWriter writer = new StreamWriter(txtfilepath, true)) //true per non eliminare e ricreare
            {
                string parameters = "Grasso (%p/V), Proteine (%p/V), Lattosio (%p/p), Caseine (%), Cellule somatiche (cell*1000/mL), Carica batterica totale (UFC*1000/mL)\n\n";
                //List<string> xlsfilepathlist = GetXlsFilePathList();
                HSSFWorkbook hssfworkbook;
                try
                {
                    using (FileStream xlsfile = new FileStream(xlsfilepath, FileMode.Open, FileAccess.Read))
                    {
                        hssfworkbook = new HSSFWorkbook(xlsfile);
                        ISheet sheet = hssfworkbook.GetSheetAt(0);
                        string producerdata = sheet.GetRow(6).GetCell(0).ToString();
                        string producerIDname = GetProducerIDName(producerdata);
                        if (IsTextFileEmpty(txtfilepath))
                        {
                            writer.WriteLine(parameters);
                        }
                        var data = GetFileData(sheet);
                        data.ForEach(line => writer.WriteLine(line));
                        writer.Close();
                    }
                    Console.WriteLine("Text file created.");
                }
                catch
                {
                    string xlsfilename = xlsfilepath.Replace($@"{projectdirpath}XLS\", "");
                    Console.WriteLine($"Create txt file and add data ERROR: {xlsfilename}.");
                }
            }
        }

        public static bool IsTextFileEmpty(string filename)
        {
            FileInfo fileinfo = new FileInfo($"{filename}");
            if (fileinfo.Length < 6) //il .Length restituisce il peso del file in bytes
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static List<string> GetFileData(ISheet sheet)
        {
            //TODO: mancano le tre date, il campione ed il codice ASL, inserire anche sui parametri prima del grasso; ordine: ASL, campione, data1, data2, data3
            List<string> datalist = new List<string>();
            Dictionary<string, int> columnDictionary = new Dictionary<string, int>();

            var haederRow = sheet.GetRow(6).Cells;
            foreach (var column in haederRow)
            {
                if (column.StringCellValue.Contains("Grasso p/v") || column.StringCellValue.Contains("Grasso (per calcolo)") || column.StringCellValue.Contains("Grasso\np/v") || column.StringCellValue.Contains("Grasso\n(per calcolo)"))
                {
                    columnDictionary.Add("Grasso", column.ColumnIndex);
                }
                else if (column.StringCellValue.Contains("Proteine p/v") || column.StringCellValue.Contains("Proteine (per calcolo") || column.StringCellValue.Contains("Grasso\np/v") || column.StringCellValue.Contains("Grasso (per calcolo)"))
                {
                    columnDictionary.Add("Proteine", column.ColumnIndex);
                }
                else if (column.StringCellValue.Contains("Lattosio"))
                {
                    columnDictionary.Add("Lattosio", column.ColumnIndex);
                }
                else if (column.StringCellValue.Contains("Caseine"))
                {
                    columnDictionary.Add("Caseine", column.ColumnIndex);
                }
                else if (column.StringCellValue.Contains("Cellule\nsomatiche"))
                {
                    columnDictionary.Add("Cellule somatiche", column.ColumnIndex);
                }
                else if (column.StringCellValue == "Carica\nBatterica\nTotale" || column.StringCellValue == "Carica Batterica Totale")
                {
                    columnDictionary.Add("Carica batterica totale", column.ColumnIndex);
                }
            }
            for (var rowindex = 8; sheet.GetRow(rowindex) != null; rowindex++)
            {
                var row = sheet.GetRow(rowindex);
                //TODO: cercare come prendere valori reali (non approssimati) da file xls
                //posso lasciare stringa, tanto i valori li prende dal file di testo
                string fat = GetData(columnDictionary, row, "Grasso");
                string protein = GetData(columnDictionary, row, "Proteine");
                string lactose = GetData(columnDictionary, row, "Lattosio");
                string casein = GetData(columnDictionary, row, "Caseine");
                string somaticcells = GetData(columnDictionary, row, "Cellule somatiche");
                string totalbacterialload = GetData(columnDictionary, row, "Carica batterica totale");
                if (fat != "" || protein != "" || lactose != "" || casein != "" || somaticcells != "" || totalbacterialload != "")
                {
                    datalist.Add($"{fat},{protein},{lactose},{casein},{somaticcells},{totalbacterialload}"); //formato csv separato da virgole
                }
            }
            return datalist;
        }

        private static string GetData(Dictionary<string, int> columnDictionary, IRow row, string column)
        {
            string value;
            try
            {
                value = row.GetCell(columnDictionary[column]).ToString().Replace("\n", " ").Trim();
            }
            catch
            {
                value = "/"; //Valore non presente nella tabella
            }
            return value;
        }
    }
}