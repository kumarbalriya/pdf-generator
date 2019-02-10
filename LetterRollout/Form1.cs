using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace LetterRollout
{
    public partial class Form1 : Form
    {
        private string pdfOutputFilePath, sourceWordFilePath, workingWordFilePath, excelFilePath;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dialogResult = wordDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                sourceWordFilePath = wordDialog.FileName;
                txtWordFileName.Text = sourceWordFilePath;
                workingWordFilePath = Path.GetDirectoryName(sourceWordFilePath);
            }
        }

        private void btnExcelFileBrowse_Click(object sender, EventArgs e)
        {
            var dialogResult = excelDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                excelFilePath = excelDialog.FileName;
                txtExcelTextbox.Text = excelFilePath;
            }
        }

        private void btnOutputFolderBrowse_Click(object sender, EventArgs e)
        {
            var dialogResult = outputFolderDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                pdfOutputFilePath = outputFolderDialog.SelectedPath;
                txtOutputFolder.Text = pdfOutputFilePath;
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(pdfOutputFilePath) ||
                string.IsNullOrEmpty(sourceWordFilePath) ||
                string.IsNullOrEmpty(workingWordFilePath) ||
                string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Please submit required data");
                return;
            }

            try
            {
                ExcelProcessor excelProcessor = new ExcelProcessor();
                WordProcessor wordp = new WordProcessor(pdfOutputFilePath, sourceWordFilePath, workingWordFilePath);
                SendMail mail = new SendMail();
                var models = excelProcessor.Main(excelFilePath);

                for (int i = 0; i < models.Count; i++)
                {
                    if (!string.IsNullOrEmpty(models[i].empId))
                    {
                        Dictionary<string, string> dict = new Dictionary<string, string>
                {
                    { "<<name>>", models[i].name },
                    { "<<empid>>", models[i].empId },
                    { "<<agc>>", models[i].agc },
                    { "<<apb>>", models[i].apb },
                    { "<<atc>>", models[i].atc },
                    { "<<basic>>", models[i].basic },
                    { "<<hra>>", models[i].hra },
                    { "<<pf>>", models[i].pf },
                    { "<<flexible>>", models[i].flexible },
                    { "<<totalearnings>>", models[i].total },
                    { "<<firstname>>", models[i].firstName },
                    { "<<title>>", models[i].title},
                    { "<<shareoptions>>", models[i].shareoptions},
                    { "<<sharegrant>>", models[i].shareGrant}
                };

                        var pdfPath = wordp.GeneratePDF(dict, models[i].empId, models[i].firstName);
                        if (!string.IsNullOrEmpty(pdfPath))
                        {
                            lstLetterGenerated.Items.Add(models[i].empId + " | " + models[i].name);
                            //mail.Main(models[i], pdfPath);
                            lstMailSent.Items.Add(models[i].empId + " | " + models[i].name);
                            Console.WriteLine((i + 1) + " Success empid: " + models[i].empId);
                        }
                        else
                        {
                            Console.Write("Something wrong processing: " + models[i].empId + " | " + models[i].name);
                        }
                    }

                }

                MessageBox.Show("Done");
                Console.WriteLine("Done");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
