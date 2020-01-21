#region librerie
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Text;
#endregion

namespace PdfToCsv
{
    class Program
    {
        #region costante globale percorso cartella
        public const string pdfdirpath = @"C:\Users\Giorgio Della Roscia\Desktop\ML\Progetti\SautinSoft\PdfToCsv\PDF"; //percorso dove si trovano gli allegati in formato pdf
        public const string txtdirpath = @"C:\Users\Giorgio Della Roscia\Desktop\ML\Progetti\SautinSoft\PdfToCsv\TXT";
        #endregion
        static void Main(string[] args)
        {
            try
            {
                List<FileInfo> files = GetFileList(); //elenco path files in files di tipo lista
                foreach (var file in files) //per ogni file contenuto in files
                {
                    #region copio i pdf e li trasformo in xls
                    SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
                    string pdfdir = file.FullName.Replace(".pdf", ""); //tolgo l'estensione ai path dei pdf e li appoggio in pdfdir
                    string xlsdir = pdfdir.Replace(@"\PDF\", @"\XLS\"); //li copio in xlsdir
                    f.OpenPdf($"{pdfdir}.pdf"); //apro i "vecchi" pdf
                    f.ToExcel($"{xlsdir}.xls"); //trasformo i nuovi file senza estensione in xls
                    string pathxls = $"{xlsdir}.xls"; //variabile d'appoggio per avere il path compatto in una sola variabile  
                    string filename = xlsdir.Replace(@"C:\Users\Giorgio Della Roscia\Desktop\ML\Progetti\SautinSoft\PdfToCsv\XLS\", "");
                    #endregion

                    #region leggo il contenuto della cella dovi si trovano i dati del produttore
                    HSSFWorkbook hssfworkbook;
                    using (FileStream excelfile = new FileStream(pathxls, FileMode.Open, FileAccess.Read))
                    {
                        hssfworkbook = new HSSFWorkbook(excelfile);
                    }
                    ISheet sheet = hssfworkbook.GetSheetAt(0);
                    string producerdata = sheet.GetRow(6).GetCell(0).ToString(); //dati del produttore (nome, ID, tipo latte)
                    #endregion 

                    #region sottostringhe produttore (ID, nome)
                    int indexA = producerdata.IndexOf('-');
                    int indexB = producerdata.IndexOf('a'); //non metto \n altrimenti prende quello dopo "Produttore:" e neanche 'L' perchè può esserci nel nome
                    string producerID = producerdata.Substring(12, indexA - 12);  //restituisce il codice del produttore
                    string producername = producerdata.Substring(indexA + 1, (indexB - 3) - indexA); //restituisce cognome e nome del produttore
                    producername = producername.Replace("\n", " "); //alcuni nomi anzichè lo spazio avevano il carattere \n
                    #endregion

                    #region creo file .txt rinominato ID e nome del produttore poi dovro inseerire i dati
                    string txtpath = txtdirpath + @"\" + producerID + "-" + producername + ".txt"; //path del nuovo file di testo
                    using (StreamWriter sw = new StreamWriter(txtpath, true)) //true per non eliminare e ricreare
                    {
                        sw.WriteLine("Grasso (%p/V); Proteine (%p/V); Lattosio (%p/p); Cellule somatiche (cell*1000/mL); Carica batterica totale (UFC*1000/mL); Caseine (%)\n");
                        sw.WriteLine("data"); //dove data rappresenta l'estratto dell'xls controllato e posto nell'ordine corretto e compreso nei giusti paramentri
                        //data deve contenere i parametri necessari ed andare a capo ogni fine riga (controllo colonna==null (anche colonna+1==null))
                        sw.Close();
                    }
                    #endregion

                    //esempio riga: n; n; n; n; n; n //con i ; anzichè , perchè sono tipo double
                    //se alcune colonne mancano sostituire con '/'
                    //ogni volta eseguire il controllo se sono presenti o meno e in quale posizione (colonna)
                }
                Console.WriteLine("processo terminato");
                Console.ReadLine();
            }
            #region catch
            catch (ArgumentException ae)
            {
                Console.WriteLine($"Argument Exception - The process failed: {ae.ToString()}.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception - The process failed: {e.ToString()}.");
            }
            Console.ReadLine();
            #endregion
        }
        #region eccezioni
        private static List<FileInfo> GetFileList()  //restituisce tutti i path dei file contenuti in dirpath come lista
        {
            DirectoryInfo di = new DirectoryInfo(pdfdirpath);
            return di.GetFiles().ToList();
        }
        #endregion
    }
}