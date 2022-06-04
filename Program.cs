using System;
using System.CodeDom;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;
using Company = SAPbobsCOM.Company;
using System.Xml;
using System.Xml.Linq;
using System.Threading;
using SAP_Helper;

namespace BINLocationAutoSelect
{
    public class Program
    {
        public static string SelectedLocatorName = "";
        public static string SelectedSeries = "";
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }


                Helpers.GlobalVar._oSAP = new SAP_Helper.SAP();
                Helpers.GlobalVar._oSAP.oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                List<string> findcols = new List<string>();

                Helpers.GlobalVar._oSAP.CreateUDF("DLN1", "AXC_BINLocation", "BIN Location", BoFieldTypes.db_Alpha, 132);

                Helpers.GlobalVar.myCompany = new Company();
                Helpers.GlobalVar.myCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Helpers.GlobalVar.oRS = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Helpers.GlobalVar.oRSExec = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Helpers.GlobalVar.oRSQuery = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Helpers.GlobalVar.oRSHdr = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Helpers.GlobalVar.oRSDtl = (Recordset)Helpers.GlobalVar.myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                Helpers.GlobalVar.oSuperUser = Helpers.GlobalVar.IsSuperUser();

                string DbUserName = Helpers.GlobalVar.myCompany.DbUserName;
                Application.SBO_Application.SetStatusBarMessage("DbUserName:" + DbUserName + ".", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                //Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }


        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes eventType)
        {
            try
            {
                switch (eventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        //Exit Add-On
                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        System.Windows.Forms.Application.Exit();
                        break;
                    default:
                        break;
                }
            }
            catch
            { }
        }
        //static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        //{
        //    //BubbleEvent = false;
        //    //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormTypeEx == "1470000007" && pVal.ItemUID == "1470000001" && BubbleEvent == true)
        //    //{
        //    //    SelectedLocatorName = "";
        //    //    Matrix _binsTbl = (Matrix)Application.SBO_Application.Forms.ActiveForm.Items.Item("1470000023").Specific;
        //    //    int totalItems = _binsTbl.RowCount;
        //    //    for (int i = 1; i <= totalItems; i++)
        //    //    {
        //    //        EditText _binCodeCol = (EditText)_binsTbl.Columns.Item("1470000001").Cells.Item(i).Specific;
        //    //        EditText _allocated = (EditText)_binsTbl.Columns.Item("1470000005").Cells.Item(i).Specific;
        //    //        if (!String.IsNullOrEmpty(_allocated?.Value ?? "") && _allocated.Value != "0")
        //    //        {
        //    //            SelectedLocatorName += String.IsNullOrEmpty(_binCodeCol) "{_binCodeCol?.Value ?? ""}; ";
        //    //        }
        //    //    }
        //    //}

        //    //if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && pVal.FormTypeEx == "140" && !String.IsNullOrEmpty(SelectedLocatorName))
        //    //{
        //    //    Matrix fmsForm = (Matrix)Application.SBO_Application.Forms.ActiveForm.Items.Item("38").Specific;
        //    //    int selrow1 = fmsForm.GetCellFocus().rowIndex;
        //    //    EditText _itemCodeCol = (EditText)fmsForm.Columns.Item("U_AXC_BINLocation").Cells.Item(selrow1).Specific;
        //    //    _itemCodeCol.Value = SelectedLocatorName;
        //    //    SelectedLocatorName = "";
        //    //}
        //}
    }
}