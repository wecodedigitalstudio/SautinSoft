//TODO: interpretazione file-->report-->

#region librerie
using iTextSharp.text;
using iTextSharp.text.pdf;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
#endregion

namespace PdfToCsv
{
    class Program
    {
        #region costante globale percorso cartella
        public const string projectdirpath = @"C:\Users\Giorgio Della Roscia\Desktop\ML\Progetti\SautinSoft\PdfToCsv\";
        #endregion
        static void Main(string[] args)
        {
            List<FileInfo> files = GetFilePathList(); 
            foreach (var file in files)
            {
                try
                {
                    List<string> listPdfSplitted = SplitPdfFileInSinglePage(file);
                    List<string> xlsList= CreateXlsFile(listPdfSplitted);
                    //string txtfilename = ExtrapolateFileName(fullxlspath);
                    //string a = CreateTxtFile(txtfilename);
                }
                catch (ArgumentException ae)
                {
                    Console.WriteLine($"Argument Exception - The process failed: {ae.ToString()}.");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Exception - The process failed: {e.ToString()}.");
                }
            }
            Console.ReadLine();
        }

        private static List<FileInfo> GetFilePathList()
        {
            DirectoryInfo di = new DirectoryInfo($@"{projectdirpath}\PDF");
            return di.GetFiles().ToList();
        }

        private static List<string> SplitPdfFileInSinglePage(FileInfo file)
        {
            List<string> fileList = new List<string>();

            string newFullName = "";

            using (PdfReader pdfreader = new PdfReader(file.FullName))
            {
                for (int pageNumber = 0; pageNumber < pdfreader.NumberOfPages; pageNumber++)
                {
                    string newName = file.Name.Replace(".pdf", "");
                    newFullName = string.Format($@"{projectdirpath}SplittedPDF\{newName}_page{pageNumber}");

                    Document document = new Document();
                    PdfCopy copy = new PdfCopy(document, new FileStream($"{newFullName}.pdf", FileMode.Create));
                    document.Open();

                    if (pageNumber < pdfreader.NumberOfPages)
                    {
                        copy.AddPage(copy.GetImportedPage(pdfreader, pageNumber+1));
                    }
                    else
                    {
                        break;
                    }
                    document.Close();
                    fileList.Add(newFullName);
                }
            }
            return fileList;
        }


        private static List<string> CreateXlsFile(List<string> listPdfFiles)
        {
            List<string> fileList = new List<string>();
            listPdfFiles.ForEach(pdfFileName =>
            {
                SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
                string xlsFullName = pdfFileName.Replace(@"\SplittedPDF\", @"\XLS\"); //li copio in xlsdir
                try
                {
                    f.OpenPdf($"{pdfFileName}.pdf"); //apro i "vecchi" pdf
                    f.ToExcel($"{xlsFullName}.xls"); //trasformo i nuovi file senza estensione in xls
                    string compactpathxls = $"{xlsFullName}.xls";
                    fileList.Add(compactpathxls);
                } catch
                {
                    Console.WriteLine($"Non è stato possibile leggere il file {pdfFileName}.");
                }
            });
            return fileList;
        }


        //    private static string ExtrapolateFileName(string fullxlspath)
        //    {
        //        HSSFWorkbook hssfworkbook;
        //        using (FileStream excelfile = new FileStream(fullxlspath, FileMode.Open, FileAccess.Read))
        //        {
        //            hssfworkbook = new HSSFWorkbook(excelfile);
        //        }
        //        ISheet sheet = hssfworkbook.GetSheetAt(0);
        //        string producerdata = sheet.GetRow(6).GetCell(0).ToString();
        //        int startindex = producerdata.IndexOf(':') + 2;
        //        int endindex = producerdata.IndexOf('a') - 2; //non metto \n altrimenti prende quello dopo "Produttore:" e neanche 'L' perchè può esserci nel nome
        //        string producerIDname = producerdata.Substring(startindex, endindex - startindex);
        //        producerIDname.Replace("\n", " "); //alcuni nomi anzichè lo spazio avevano il carattere \n
        //        return producerIDname;
        //    }
        //    private static string CreateTxtFile(string txtfilename)
        //    {
        //        string txtfilepath = $@"{projectdirpath}\TXT\{txtfilename}.txt";
        //        string parameters = "Grasso (%p/V); Proteine (%p/V); Lattosio (%p/p); Cellule somatiche (cell*1000/mL); Carica batterica totale (UFC*1000/mL); Caseine (%)\n";
        //        using (StreamWriter sw1 = new StreamWriter(txtfilepath, true)) //true per non eliminare e ricreare
        //        {
        //            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
        //            ISheet sheet = hssfworkbook.GetSheetAt(0);
        //            string[] lines = File.ReadAllLines(txtfilepath);
        //            foreach (string line in lines)
        //            {
        //                StreamReader sr1 = new StreamReader(txtfilepath);
        //                bool comparisonresult = line.Equals(parameters);
        //                if (sheet.GetRow(6).GetCell(0) != null)
        //                {
        //                    if (line.Count != 0)
        //                    {
        //                        sw1.WriteLine(data);
        //                    }
        //                    else
        //                    {
        //                        sw1.WriteLine(parameters);
        //                    }
        //                }
        //            }
        //            sw1.Close();
        //        }
        //        return ;
        //    }
    }
    }