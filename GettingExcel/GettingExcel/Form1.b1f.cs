using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Globalization;

namespace GettingExcel
{
    [FormAttribute("GettingExcel.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
           
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;

        object missing = Type.Missing;
   
    
        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbobsCOM.Company ocompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Recordset rset = (SAPbobsCOM.Recordset)ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            DateTime startdate = Convert.ToDateTime(objBridge.Format_StringToDate(EditText0.Value).Fields.Item(0).Value);
            DateTime enddate = Convert.ToDateTime(objBridge.Format_StringToDate(EditText1.Value).Fields.Item(0).Value);
            Excel.Application oXL = new Excel.Application();
            oXL.Visible = false; 
            Excel.Workbook oWB = oXL.Workbooks.Add(missing);
            
          
            

            string callprocedure = $"call \"Excel_Get_JournalEntry_Header\"('{startdate.ToString("yyyy-MM-dd")}' , '{enddate.ToString("yyyy-MM-dd")}')";


            StaticText0.Caption = startdate.ToString("yyyy-MM-dd");

            rset.DoQuery(callprocedure);
            if (rset.RecordCount > 0)
            {
                

                Excel.Worksheet oSheet = oWB.ActiveSheet as Excel.Worksheet;
                oSheet.Name = "defterMain";
                oSheet.Cells[1, 1] = "vkn";
                oSheet.Cells[1, 2] = "Period_start";
                oSheet.Cells[1, 3] = "Period_end";
                oSheet.Cells[1, 4] = "Sube_kodu";
                oSheet.Columns.AutoFit();
                var headcolorsh1 = oSheet.Range[
                    oSheet.Cells[1, 1],
                    oSheet.Cells[1, 4]];
                headcolorsh1.Interior.Color = Excel.XlRgbColor.rgbDarkGray;
                oSheet.Columns.AutoFit();


                Excel.Worksheet oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing)
                                as Excel.Worksheet;
                oSheet2.Name = "entryHeader";
                int b = rset.Fields.Count;
                var range1 = oSheet2.Cells[1, b];
               
                for (var i = 0; i < rset.Fields.Count; i++)
                {
                    rset.MoveFirst();
                    oSheet2.Cells[1, i + 1] = rset.Fields.Item(i).Description;
                    for (var j = 0; j < rset.RecordCount; j++)
                    {
                        oSheet2.Cells[j + 2, i + 1] = rset.Fields.Item(i).Value;
                        rset.MoveNext();
                    }                   

                }

                var headcolorsh2 = oSheet2.Range[
                       oSheet2.Cells[1, 1],
                       oSheet2.Cells[1, b]];
                headcolorsh2.Interior.Color = Excel.XlRgbColor.rgbDarkGray;
                oSheet2.Columns.AutoFit();
                
               
            }
            rset = (SAPbobsCOM.Recordset)ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string callprocedure2 = $"call \"Excel_Get_JournalEntry_Detail\"('{startdate.ToString("yyyy-MM-dd")}' , '{enddate.ToString("yyyy-MM-dd")}')";
            rset.DoQuery(callprocedure2);
            if (rset.RecordCount > 0)
            {
                Excel.Worksheet oSheet3 = oWB.Sheets.Add(missing, missing, 1, missing)
                              as Excel.Worksheet;
                oSheet3.Name = "entryDetail";
                int s = rset.Fields.Count;
                for (var i = 0; i < rset.Fields.Count; i++)
                {
                    rset.MoveFirst();
                    oSheet3.Cells[1, i + 1] = rset.Fields.Item(i).Description;
                    for (var j = 0; j < rset.RecordCount; j++)
                    {
                       
                        oSheet3.Cells[j + 2, i + 1] = rset.Fields.Item(i).Value;
                        rset.MoveNext();
                       
                    }
                  

                }
                var headcolorsh3 = oSheet3.Range[
                    oSheet3.Cells[1, 1],
                    oSheet3.Cells[1, s]];
                headcolorsh3.Interior.Color = Excel.XlRgbColor.rgbDarkGray;
                oSheet3.Columns.AutoFit();
                

            }



           

           

           

            var datetoday = (DateTime.Today);

           
            
            
            string fileName = "work" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xls";
            oWB.SaveAs(@"d:\" + fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,
                        missing, missing, missing, missing,
                        Excel.XlSaveAsAccessMode.xlNoChange,
                        missing, missing, missing, missing, missing);
            oWB.Close(missing, missing, missing);
            oXL.UserControl = true;
            oXL.Quit();







        }

        private SAPbouiCOM.StaticText StaticText0;
    }






    }
