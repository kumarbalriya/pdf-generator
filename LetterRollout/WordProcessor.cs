using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace LetterRollout
{
    public class WordProcessor
    {
        Application application;
        string pdfOutputDirectory;
        string sourceWordFilePath;
        string workingWordFilePath;

        public WordProcessor(string _pdfOutputDirectory, string _sourceWordFilePath, string _workingWordFilePath)
        {
            pdfOutputDirectory = _pdfOutputDirectory;
            sourceWordFilePath = _sourceWordFilePath;
            workingWordFilePath = _workingWordFilePath + @"\WorkingWord1.docx";

            application = new Application();
            application.Visible = false;
        }

        public string GeneratePDF(Dictionary<string, string> dict, string empId, string firstName)
        {
            object missing = Type.Missing;
            //object pdfOutputFilePath = ConfigurationSettings.AppSettings["pdfOutputFilePath"] + firstName + "_" + empId + ".pdf";
            object pdfOutputFilePath = pdfOutputDirectory + @"\" + firstName + "_" + empId + ".pdf";
            object pdfFormat = WdSaveFormat.wdFormatPDF;
            //string sourceWordFilePath = ConfigurationSettings.AppSettings["sourceWordFilePath"];
            //string workingWordFilePath = ConfigurationSettings.AppSettings["workingWordFilePath"];
            bool success = false;

            File.Copy(sourceWordFilePath, workingWordFilePath);

            Document doc = application.Documents.Open(workingWordFilePath);
            
            try
            {
                foreach (var item in dict)
                {
                    Find findObject = application.Selection.Find;
                    findObject.ClearFormatting();
                    findObject.Text = item.Key;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = item.Value;

                    object replaceAll = WdReplace.wdReplaceAll;
                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }

                doc.SaveAs2(ref pdfOutputFilePath, ref pdfFormat);
                success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                doc.Close();
                File.Delete(workingWordFilePath);
            }

            return success == true ? pdfOutputFilePath.ToString() : "";
        }
    }
}
