using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace LetterRollout
{
    public class ExcelProcessor
    {
        public List<Model> Main(string excelFilePath)
        {
            Application application = new Application();
            //var excelFilePath = ConfigurationSettings.AppSettings["excelFilePath"];
            Workbook workbook = application.Workbooks.Open(excelFilePath);
            List<Model> models = new List<Model>();

            try
            {

                Worksheet workSheet = workbook.Worksheets[1];
                var usedRange = workSheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    Model m = new Model() { };
                    m.emailId = (workSheet.Cells[i, 1]).Value;
                    m.name = (workSheet.Cells[i, 2]).Value;
                    m.empId = Convert.ToString((workSheet.Cells[i, 3]).Value);
                    m.agc = Convert.ToString((workSheet.Cells[i, 4]).Value);
                    m.apb = Convert.ToString((workSheet.Cells[i, 5]).Value);
                    m.atc = Convert.ToString((workSheet.Cells[i, 6]).Value);
                    m.basic = Convert.ToString((workSheet.Cells[i, 7]).Value);
                    m.hra = Convert.ToString((workSheet.Cells[i, 8]).Value);
                    m.pf = Convert.ToString((workSheet.Cells[i, 9]).Value);
                    m.flexible = Convert.ToString((workSheet.Cells[i, 10]).Value);
                    m.total = Convert.ToString((workSheet.Cells[i, 11]).Value);
                    m.firstName = (workSheet.Cells[i, 12]).Value;
                    m.title = (workSheet.Cells[i, 13]).Value;
                    m.shareoptions = (workSheet.Cells[i, 14]).Value;
                    m.shareGrant = Convert.ToString((workSheet.Cells[i, 15]).Value);
                    m.cc = (workSheet.Cells[i, 16]).Value;

                    models.Add(m);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                workbook.Close();
            }

            return models;
        }
    }
}
