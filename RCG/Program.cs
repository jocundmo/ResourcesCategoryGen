using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Xml;
using System.IO;

namespace RCG
{
    // Bug1: Movie.HD as existing xml, but disabled it roll in Game as enabled xml. no response.
    class Program
    {
        private static void GenerateExcel()
        {
            Excel.Application excel = new Excel.Application();
            int rowIndex = 1;
            int colIndex = 0;

            excel.Application.Workbooks.Add(true);


            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Gender", typeof(string));

            DataRow dr = table.NewRow();
            dr["Name"] = "abc";
            dr["Gender"] = "M";
            table.Rows.Add(dr);

            foreach (DataColumn col in table.Columns)
            {
                colIndex++;
                excel.Cells[1, colIndex] = col.ColumnName;
                //Range r = excel.Cells[1, colIndex] as Range;
            }

            foreach (DataRow row in table.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in table.Columns)
                {
                    colIndex++;
                    excel.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
                }
            }
            excel.Visible = false;

            excel.ActiveWorkbook.SaveAs("C:\\abc2.xls");
            //excel.ActiveWorkbook.SaveAs("C:\\A.XLS", Excel.XlFileFormat.xlExcel9795, null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);



            //wkbNew.SaveAs strBookName


            //excel.Save(strExcelFileName);
            excel.Quit();
            excel = null;

            GC.Collect();//垃圾回收
        }

        static MessageLogger logger = new MessageLogger(string.Format("RCG_log_{0}.txt", DateTime.Now.ToString("yyyyMMdd-HHmmss")));

        // RCG.exe -forceEnable=高清电影,游戏
        // RCG.exe HD_Movies_Mappings.xml true
        static void Main(string[] args)
        {
            //args = new string[2];
            //args[0] = "<Sheet name='游戏' enabled='true' mode='append' maxRowCount='1000'></Sheet>";
            //args[1] = "true";
            string configFileName = "Mappings.xml";
            string hijackConfigurationParameter = string.Empty;
            bool hijackOnlyAttribute = false;
            if (args != null && args.Length > 1)
            {
                logger.LogMessage(string.Format("Arguments print: {0}", args[0]));
                logger.LogMessage(string.Format("Arguments print: {0}", args[1]));

                //logger.LogMessage("Reading excel row: " + e.RowProcessing[int.Parse(e.KeyMessage)].ToString() + "...");
                hijackConfigurationParameter = args[0].Trim();
                hijackOnlyAttribute = bool.Parse(args[1].Trim());
            }

            GenProcessor gp = new GenProcessor();
            gp.OnGeneralMessageException += new EventHandler<GeneralMessageEventArgs>(gp_OnGeneralMessageException);
            gp.OnHandlableException += new EventHandler<HandlableExceptionEventArgs>(gp_OnHandlableException);
            gp.OnReadingMetadata += new EventHandler<DataRowEventArgs>(gp_OnReadingMetadata);
            gp.OnReadingExcelRow += new EventHandler<DataRowEventArgs>(gp_OnReadingExcelRow);
            gp.OnProcessingMetadata += new EventHandler<DataRowEventArgs>(gp_OnProcessingMetadata);
            //gp.OnFormatting += new EventHandler<DataRowEventArgs>(gp_OnFormatting);
            //gp.OnSettingRowMode += new EventHandler<DataRowEventArgs>(gp_OnSettingRowMode);
            //gp.OnFiltering += new EventHandler<DataRowEventArgs>(gp_OnFiltering);
            gp.OnWritingDataRow += new EventHandler<DataRowEventArgs>(gp_OnWritingDataRow);

            try
            {
                logger.LogMessage("[Step 1/5] Reading configuration...");
                gp.ReadConfiguration(configFileName, hijackConfigurationParameter, hijackOnlyAttribute);
                
                logger.LogMessage("[Step 2/5] Generating metadata...");
                gp.GenerateMetadata();

                logger.LogMessage("[Step 3/5] Reading excel...");
                //gp.OutputTemporaryFiles("rcg_temp.xml");
                gp.ReadPreviousMetadata("rcg_post_temp.xml");

                logger.LogMessage("[Step 4/5] Processing metadata table...");
                gp.ProcessMetadataTable();

                gp.OutputTemporaryFiles("rcg_post_temp.xml");

                logger.LogMessage("[Step 5/5] Generating excel...");
                gp.RefreshExcel();

                logger.LogMessage("Done successfully!!!");

            }
            catch (Exception ex)
            {
                logger.LogMessage("Error found: " + ex.ToString());
            }
            finally
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }

            //回收垃圾
            //public void KillWordProcess() 
            //{ 
            //int ProceedingCount = 0; 
            //try 
            //{ 
            //System.Diagnostics.Process [] ProceddingCon = System.Diagnostics.Process.GetProcesses(); 
            //foreach(System.Diagnostics.Process IsProcedding in ProceddingCon) 
            //{ 
            //if(IsProcedding.ProcessName.ToUpper() == "WINWORD") 
            //{ 
            //ProceedingCount += 1; 
            //IsProcedding.Kill(); 
            //} 
            //} 
            //} 
            //catch(System.Exception err) 
            //{ 
            //MessageBox.Show(err.Message + "\r" +"(" + err.Source + ")" + "\r" + err.StackTrace); 
            //} 
            //} 
            //#endregion


        }

        static void gp_OnGeneralMessageException(object sender, GeneralMessageEventArgs e)
        {
            string message = e.Message;
            logger.LogMessage(message);
        }

        static void gp_OnHandlableException(object sender, HandlableExceptionEventArgs e)
        {
            string message = string.IsNullOrEmpty(e.KeyMessage) ? e.HandlableException.ToString() : e.KeyMessage;
            logger.LogMessage("Exception is handled: " + message);
        }

        static void gp_OnReadingExcelRow(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Reading excel row: " + e.RowProcessing[int.Parse(e.KeyMessage)].ToString() + "...");
        }

        static void gp_OnSettingRowMode(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Setting row mode:  " + e.RowProcessing[e.KeyMessage].ToString() + "...");
        }

        static void gp_OnFormatting(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Formatting:        " + e.RowProcessing[e.KeyMessage].ToString() + "...");
        }

        static void gp_OnWritingDataRow(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Writing data row:  " + e.RowProcessing[Constants.COLUMN_RowMode] + " " + e.RowProcessing[e.KeyMessage].ToString() + "...");
        }

        static void gp_OnProcessingMetadata(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Processing:        " + e.RowProcessing[e.KeyMessage].ToString() + "...");
        }

        static void gp_OnFiltering(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Filtering:         " + e.RowProcessing[e.KeyMessage].ToString() + "...");
        }

        static void gp_OnReadingMetadata(object sender, DataRowEventArgs e)
        {
            logger.LogMessage("Reading meatadata: " + e.KeyMessage + "...");
        }
    }
}
