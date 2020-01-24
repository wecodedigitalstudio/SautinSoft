//TODO: interpretazione file-->report-->variabile di diverse tipologie
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
            foreach (var pdffile in pdffiles)
            {
                try
                {
                    List<string> listpdfsplitted = SplitPdfFileInSinglePage(pdffile);
                    CreateTxtFilesWithoutHavingXls(listpdfsplitted); //usare se devo creare xlsfile
                    CreateTxtFilesHavingXls();//usare se ho già creato i miei xls file
                }

                catch (ArgumentException)
                {
                    Console.WriteLine($"AE - Errore nella trasformazione del file: {pdffile.Name}.");
                    //File.Delete(pdffile.FullName);
                }
                catch (Exception)
                {
                    Console.WriteLine($"E - Errore nella trasformazione del file: {pdffile.Name}.");
                    //File.Delete(pdffile.FullName);
                }
            }
            Console.ReadLine();
        }
        #region
        private static List<FileInfo> GetPdfFilePathList(string folder)
        {
            DirectoryInfo di = new DirectoryInfo($@"{projectdirpath}\{folder}");
            return di.GetFiles().ToList();
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
                    else
                    {
                        break;
                    }
                    document.Close();
                    filelist.Add(newfullname);
                }
            }
            return filelist;
        }

        private static List<string> GetXlsFilePathList()
        {
            List<string> xlsfilelist = new List<string>();
            List<FileInfo> xlsfiles = GetPdfFilePathList("XLS");
            foreach (var xlsfile in xlsfiles) //lascio il ciclo per visualizzare un eventuale errore a schermo 
            {
                try
                {
                    xlsfilelist.Add($"{xlsfile}");
                }
                catch
                {
                    Console.WriteLine($"Errore nel restituire il percorso del file xls: {xlsfile.Name}.");
                }
            }
            return xlsfilelist;
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
                    Console.WriteLine($"Non è stato possibile leggere il file {pdfFileName}.");
                }
            });
            return fileList;
        }

        private static void CreateTxtFilesHavingXls()
        {
            List<string> xlsfilelist = GetXlsFilePathList();
            foreach (string xlsfile in xlsfilelist)
            {
                string txtfilename = ExtrapolateFileName(xlsfile);
                if (txtfilename != null && txtfilename.Contains('-'))
                {
                    CreateTxtFile(txtfilename);
                }
            }
        }

        private static void CreateTxtFilesWithoutHavingXls(List<string> listpdfsplitted)
        {
            List<string> xlsfiles = CreateXlsFile(listpdfsplitted);
            foreach (string xlsfile in xlsfiles)
            {
                string txtfilename = ExtrapolateFileName(xlsfile);
                if (txtfilename != null && txtfilename.Contains('-'))
                {
                    CreateTxtFile(txtfilename);
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
            string producerIDname = "";
            if (producerdata != "")
            {
                RegexOptions options = RegexOptions.None;
                Regex regex = new Regex("[ ]{2,}", options);
                int startindex = producerdata.IndexOf(':')+2;
                int endindex = producerdata.IndexOf('a')-2; //non metto \n altrimenti prende quello dopo "Produttore:" e neanche 'L' perchè può esserci nel nome
                producerIDname = producerdata.Substring(startindex, endindex - startindex);
                producerIDname = regex.Replace(producerIDname.Replace("\n", " "), " ");
            }
            return producerIDname;
        }
        #endregion
        private static void CreateTxtFile(string txtfilename)
        {

            string txtfilepath = $@"{projectdirpath}\TXT\{txtfilename}.txt";
            using (StreamWriter writer = new StreamWriter(txtfilepath, true)) //true per non eliminare e ricreare
            {
                string parameters = "Grasso (%p/V); Proteine (%p/V); Lattosio (%p/p); Cellule somatiche (cell*1000/mL); Carica batterica totale (UFC*1000/mL); Caseine (%)\n";
                List<string> xlsfilepathlist = GetXlsFilePathList();
                HSSFWorkbook hssfworkbook;
                foreach (var xlsfilepath in xlsfilepathlist)
                {
                    try
                    {
                        using (FileStream xlsfile = new FileStream(xlsfilepath, FileMode.Open, FileAccess.Read))
                        {
                            hssfworkbook = new HSSFWorkbook(xlsfile);
                        }
                        ISheet sheet = hssfworkbook.GetSheetAt(0);
                        string headerline = File.ReadAllLines(txtfilepath)[0];
                        StreamReader reader = new StreamReader(txtfilepath);
                        string producer = sheet.GetRow(6).GetCell(0).ToString();
                        if (headerline != parameters)
                        {
                            writer.WriteLine(parameters);
                        }
                        else
                        {
                            //writer.WriteLine($@"\n{data});
                        }
                        writer.Close();
                    }
                    catch
                    {
                        Console.WriteLine($"Errore relativo al creare il file di testo ed aggiungere i dati. Nome del file: {xlsfilepath}."); //TODO: attualmente stampa l'intero path, solo il nome
                    }
                }
            }
        }
    }
}