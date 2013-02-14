using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.IO;
using System.Globalization;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using CommandLine.Utility;


namespace BDOC.MigrationTools
{
    class RenameWordStyles
    {
        public static Stack<string> docFiles = new Stack<string>();       
        private Logger logger;
        static ReaderWriterLockSlim _rw = new ReaderWriterLockSlim();
        
        public RenameWordStyles() {
            logger = new Logger("LogChangeStylesV5.txt");            
        }

        private string RemoveDiacritics(string stIn)
        {
            string stFormD = stIn.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            for (int ich = 0; ich < stFormD.Length; ich++)
            {
                UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(stFormD[ich]);
                if (uc != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(stFormD[ich]);
                }
            }

            return (Regex.Replace(sb.ToString().Normalize(NormalizationForm.FormC), "[^A-Za-z0-9]", ""));
        }

        private bool findFiles(string dir) {
            DirectoryInfo di = new DirectoryInfo(dir);
            FileInfo[] rgFiles = di.GetFiles("*.doc");
            foreach (FileInfo fi in rgFiles)
            {
                docFiles.Push(fi.FullName);
                logger.Log("File " + fi.Name + " Added.");
            }
            return true;
        }

      
        public bool startModification(string[] args)
        {
            Arguments CommandLine = new Arguments(args);

            findFiles(CommandLine["dir"]);
            if (CommandLine["single"] != null)
            {
                applyFileModification(args);
            }
            else
            {

                Thread thread1 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread2 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread3 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread4 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread5 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread6 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread7 = new Thread(new ParameterizedThreadStart(applyFileModification));
                Thread thread8 = new Thread(new ParameterizedThreadStart(applyFileModification));

                thread1.Start(args);
                thread2.Start(args);
                thread3.Start(args);
                thread4.Start(args);
                thread5.Start(args);
                thread6.Start(args);
                thread7.Start(args);
                thread8.Start(args);
            }
                    

            return true;
        }

        //Generic Method
        private void applyFileModification(object  parameters)
        {
            //popping out from the file stack
            string file;
            Document aDoc;
            int saveFile = 0;

            while (docFiles.Count > 0)
            {
                Application WordAppParallel = new Application();
                lock (docFiles)
                {
                    if (docFiles.Count > 0)
                    {
                        file = docFiles.Pop();
                        System.Console.WriteLine(file + " popped out. " + docFiles.Count + " left.");
                    }
                    else
                    {
                        ((_Application)WordAppParallel).Quit();
                        return;
                    }
                }
                logger.Log(file + "|ProcessFile|Process of File " + file + " started.");
                //--------------------

                
                // set the file name from the open file dialog
                object fileName = file;
                object readOnly = false;
                object isVisible = true;
                object OpenAndRepair = true;
                object NoEncodingDialog = true;
                // Here is the way to handle parameters you don't care about in .NET
                object missing = System.Reflection.Missing.Value;
                // Make word visible, so you can see what's happening
                WordAppParallel.Visible = false;
                // Open the document that was chosen by the dialog
                try
                {
                    aDoc = WordAppParallel.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref NoEncodingDialog, ref missing);
                    // Activate the document so it shows up in front
                    aDoc.Activate();

                    // Treatment ------------------------------

                    Arguments CommandLine = new Arguments((string[])parameters);
                   
                       
                       
                                
                        if (CommandLine["l"]!= null){                       
                            //ListTABS
                            saveFile = saveFile + listTabsV5(ref aDoc);
                        }
                                 
                        if (CommandLine["f"]!= null){     
                            //ListFonts
                            saveFile = saveFile + listFontsV5(ref aDoc);
                        }

                        if (CommandLine["m"] != null)
                        {
                            //ModifyStyles
                            saveFile = saveFile + modifyStylesV5(ref aDoc);
                        }
                        if (CommandLine["t"]!= null){  
                            //modifyTABV5
                            saveFile = saveFile + modifyTABLESV5(ref aDoc);
                        }

                        if (CommandLine["tw"] != null)
                        {
                            //modifyTABV5
                            saveFile = saveFile + modifyTABLESPreferredWidthV5(ref aDoc);
                        }

                        if (CommandLine["e"] != null){
                            //modifyTABV5
                            saveFile = saveFile + modifyInterligneV5(ref aDoc, CommandLine["e"], CommandLine["i"]);
                        }

                        
                        
                     
                    //End treatment-----------------------------

                    //Save File----
                     if (saveFile != 0) {
                         writeFile(ref aDoc);
                     }


                    logger.Log(file + "|EndProcessFile|Process of File " + file + " finished.");
                }
                catch (Exception ex)
                {
                    logger.LogFatal(file + "|ProcessFileWithErrors|Process of File " + file + " Crashed. Thread " + Thread.CurrentThread.ManagedThreadId + "|" + ex.Message);

                }
                finally
                {
                    ((_Application)WordAppParallel).Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(WordAppParallel);
                }

            }


        }

        private int writeFile(ref Document aDoc)
        {
            _rw.EnterWriteLock();
            aDoc.Save();
            aDoc.Close();
            _rw.ExitWriteLock();
            return 1;
        }

        private int modifyStylesV5(ref Document aDoc)
        {
            foreach (Style newStyles in aDoc.Styles)
            {
                if (newStyles.BuiltIn == false)
                {
                    logger.Log(aDoc.Name + "|Style| " + newStyles.NameLocal + " found.");
                    string newName;
                    newName = newStyles.NameLocal;
                    newName = RemoveDiacritics(newName);
                    newStyles.NameLocal = newName;
                    logger.Log(aDoc.Name + "|NewStyle| " + newStyles.NameLocal + " renamed.");
                }
            }
            return 1;
        }

        private int listTabsV5(ref Document aDoc)
        {
            bool result = false;
            char findText = Convert.ToChar(9);
            Range range = aDoc.Content;
            Find find = range.Find;
            // Here is the way to handle parameters you don't care about in .NET
            object missing = System.Reflection.Missing.Value;
            find.Text = findText.ToString();
            find.ClearFormatting();
            result = find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            if (result == true)
            {
                logger.Log(aDoc.Name + "|TABFOUND|" + aDoc.Name);
            }
            else {
                logger.Log(aDoc.Name + "|NOTABFOUND|" + aDoc.Name);
            }
            return 0;
        }

        private int listFontsV5(ref Document aDoc)
        {
            foreach (Style newStyles in aDoc.Styles)
            {
                if (newStyles.BuiltIn == false)
                    logger.Log(aDoc.Name + "|UserStyle| " + newStyles.NameLocal + "|Font|" + newStyles.Font.Name);
                else
                    logger.Log(aDoc.Name + "|BuiltStyle| " + newStyles.NameLocal + "|Font|" + newStyles.Font.Name);

            }
            return 0;
        }

        private int modifyTABLESV5(ref Document aDoc)
        {

            Tables range = aDoc.Content.Tables;

            foreach (Table tableNew in range)
            {
                logger.Log(aDoc.Name + "|Tablefound| found.");
                //autofitting
                tableNew.AllowAutoFit = false;
                //indentation a gauche 0
                tableNew.Rows.LeftIndent = 0;
               
                logger.Log(aDoc.Name + "|TableLeftIndent| 0 indent.");
                logger.Log(aDoc.Name + "|TableAutiFit| Autofit in False.");
            }
            return 1;
        }

        private int modifyTABLESPreferredWidthV5(ref Document aDoc)
        {

            Tables range = aDoc.Content.Tables;

            foreach (Table tableNew in range)
            {
                logger.Log(aDoc.Name + "|Tablefound| found.");
                
                //colonne taille
                tableNew.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                logger.Log(aDoc.Name + "|PreferredWidthType| pourcentage.");

            }
            return 1;
        }
        private int modifyInterligneV5(ref Document aDoc, string type, string lineSpace)
        {
            if (lineSpace == null || lineSpace=="true")
            {
                lineSpace = "10";
            }
            switch (type){
                case "1":
                    aDoc.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;                   
                    logger.Log(aDoc.Name + "|modifyInterligne| LineSpaceSingle" + lineSpace);            
                    break;
                case "2":
                    aDoc.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;                    
                    logger.Log(aDoc.Name + "|modifyInterligne| LineSpace1pt5");
                    break;
                case "3":
                    aDoc.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;                    
                    logger.Log(aDoc.Name + "|modifyInterligne| LineSpaceDouble " );
                    break;
                case "4":
                    aDoc.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                    aDoc.Paragraphs.LineSpacing = int.Parse(lineSpace);
                    logger.Log(aDoc.Name + "|modifyInterligne| LineSpaceatLeast = " + lineSpace);
                    break;
                case "5":
                    aDoc.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                    aDoc.Paragraphs.LineSpacing = int.Parse(lineSpace);
                    logger.Log(aDoc.Name + "|modifyInterligne| LineSpaceExactly = " + lineSpace);
                    break;
                case "6":
                    aDoc.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                    aDoc.Paragraphs.LineSpacing = int.Parse(lineSpace);
                    logger.Log(aDoc.Name + "|modifyInterligne| LineSpaceMultiple = " + lineSpace);
                    break;
                default:
                    return 0;
            }            
           
            return 1;
        }

    }     
}
