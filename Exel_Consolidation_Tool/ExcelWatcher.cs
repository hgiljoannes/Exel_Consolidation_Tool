using Exel_Consolidation_Tool.ExcelProcess;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exel_Consolidation_Tool
{
    public partial class ExcelWatcher : Form
    {

        public static string pathE = "";
    
      
        string subdirNotApplicable = "";
        string subdirProcessed = "";
        public static ExcelWatcher formInstance;
        public TextBox txtFlog;
        string files = "";
        copyExcelSheets sES = new copyExcelSheets();
        killExcel kE = new killExcel();



        public ExcelWatcher()
        {
            InitializeComponent();
            formInstance = this;
            txtFlog = txtLog;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog diag = new FolderBrowserDialog();
            createExcelMasterFile excelFile = new createExcelMasterFile();
            if (diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtPath.Text = diag.SelectedPath;
                pathE = diag.SelectedPath;
                fileSystemWatcher1.Path = pathE;
                subdirNotApplicable = pathE + @"\Not Applicable";
                subdirProcessed = pathE + @"\Processed";
                txtLog.Text += "Starting Monitoring..." + Environment.NewLine;
                if (!Directory.Exists(subdirNotApplicable))
                    Directory.CreateDirectory(subdirNotApplicable);

                if (!Directory.Exists(subdirProcessed))
                    Directory.CreateDirectory(subdirProcessed);

             
                    excelFile.createFile(pathE);
               
                
               
               
                
            }
            else
            {
                txtPath.Text = "You didn't select the folder!";
            }


        }

        #region GetFiles
        public void GetFiles()
        {

            var lst = Directory.GetFiles(pathE).Where(name => !name.EndsWith("ExcelRoot.xlsx") && !name.Contains("~$"));

            foreach (var sfile in lst)
            {
                if (sfile.Contains(".xlsx"))
                {
                    files = sfile + " -> Processed " + Environment.NewLine;
                    
                    txtLog.Invoke((Action)delegate
                    {
                        txtLog.Text += files;
                    });
 
                    sES.Copy(sfile, pathE);
                   
                   
                    if (!Directory.Exists(subdirProcessed))
                    {
                        Directory.CreateDirectory(subdirProcessed);

                        foreach (var file in new DirectoryInfo(pathE).GetFiles().Where(name => name.FullName == sfile))
                        {
                            string pathMove = $@"{subdirProcessed}\{file.Name}";
                            try
                            {
                                file.MoveTo(pathMove);
                            }
                            catch (Exception)
                            {
                                file.MoveTo(pathMove, true);
                            }
                        }

                    }
                    else
                    {

                        foreach (var file in new DirectoryInfo(pathE).GetFiles().Where(name => name.FullName == sfile))
                        {
                            string pathMove = $@"{subdirProcessed}\{file.Name}";
                            try
                            {
                                file.MoveTo(pathMove);
                            }
                            catch (Exception)
                            {
                                file.MoveTo(pathMove, true);
                            }

                        }
                    }
              

                
                   


                }
                else
                {
                    files = sfile + " -> Not Applicable " + Environment.NewLine;
                    txtLog.Invoke((Action)delegate
                    {
                        txtLog.Text += files;
                    });


                    // If directory does not exist, create it. 
                    if (!Directory.Exists(subdirNotApplicable))
                    {
                        Directory.CreateDirectory(subdirNotApplicable);

                        foreach (var file in new DirectoryInfo(pathE).GetFiles().Where(name => name.FullName == sfile))
                        {
                            string pathMove = $@"{subdirNotApplicable}\{file.Name}";

                            try
                            {
                                file.MoveTo(pathMove);
                            }
                            catch (Exception)
                            {
                                file.MoveTo(pathMove, true);

                            }
                        }

                    }
                    else
                    {

                        foreach (var file in new DirectoryInfo(pathE).GetFiles().Where(name => name.FullName == sfile))
                        {

                            string pathMove = $@"{subdirNotApplicable}\{file.Name}";

                            try
                            {
                                file.MoveTo(pathMove);
                            }
                            catch (Exception)
                            {
                                file.MoveTo(pathMove, true);

                            }

                        }
                    }
                }
            }

          

        }

        #endregion



        private void Form1_Load(object sender, EventArgs e)
        {
         
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

       

        private void fileSystemWatcher1_Created(object sender, FileSystemEventArgs e)
        {
            GetFiles();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            kE.KillExcel();
            this.Close();
        }
    }
}
