using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace BIN_Location_Auto_Select_Addon.Forms
{
    [FormAttribute("140", "Forms/frmDeliveryOrder.b1f")]
    class frmDeliveryOrder : SystemFormBase
    {
        #region SAP Objects
        public SAPbouiCOM.ButtonCombo btTrans { get; set; }
        public SAPbouiCOM.ComboBox cbSeries { get; set; }
        public SAPbouiCOM.EditText edDocNum { get; set; }
        public SAPbouiCOM.Matrix mxContents { get; set; }
        public SAPbouiCOM.Column mxColBin { get; set; }
        public SAPbouiCOM.Column mxAxcColBin { get; set; }
        public SAPbouiCOM.EditText mxAxcBin { get; set; }
        public SAPbouiCOM.Button btnRefresh { get; set; }
        public SAPbouiCOM.Button btnCancel { get; set; }
        //public SAPbouiCOM.LinkedButton linkedBtnBinCol { get; set; }
        #endregion SAP Objects

        #region Variables
        public static string Series { get; set; }
        public static string DocNum { get; set; }
        public static string BINLocation { get; set; }
        public static bool BinNotExist { get; set; }
        public static string mxRow { get; set; }
        public static int incErr { get; set; }
        public static int incSuccess { get; set; }
        public static bool Success { get; set; }
        public string frmType { get; set; }
        public int frmCnt { get; set; }
        public string bincod { get; set; }
        public string binqty { get; set; }
        public string oldbincod { get; set; }
        public string msgProcess { get; set; }
        #endregion Variables
        public frmDeliveryOrder()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.cbSeries = ((SAPbouiCOM.ComboBox)(this.GetItem("88").Specific));
            this.edDocNum = ((SAPbouiCOM.EditText)(this.GetItem("8").Specific));
            this.mxContents = ((SAPbouiCOM.Matrix)(this.GetItem("38").Specific));
            this.mxColBin = (SAPbouiCOM.Column)(this.mxContents.Columns.Item("1470002149")); //BinLoc
            this.mxAxcColBin = (SAPbouiCOM.Column)(this.mxContents.Columns.Item("U_AXC_BINLocation"));
            this.btnRefresh = ((SAPbouiCOM.Button)(this.GetItem("refBtnDlv").Specific));
            this.btTrans = ((SAPbouiCOM.ButtonCombo)(this.GetItem("1").Specific));
            this.btnCancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.btnRefresh.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnRefreshClickAfter);
            this.btTrans.ClickBefore += new SAPbouiCOM._IButtonComboEvents_ClickBeforeEventHandler(this.btTrans_ClickBefore);
            this.btTrans.ClickAfter += new SAPbouiCOM._IButtonComboEvents_ClickAfterEventHandler(this.btTrans_ClickAfter);
            this.btTrans.PressedBefore += new SAPbouiCOM._IButtonComboEvents_PressedBeforeEventHandler(this.btTrans_PressedBefore);
            //this.linkedBtnBinCol.ClickAfter += new SAPbouiCOM._ILinkedButtonEvents_ClickAfterEventHandler(this.linkbtBinCol_ClickAfter);
            this.mxContents.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.mxContents_GotFocusAfter);
            this.mxContents.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.mxContents_LostFocusAfter);
            this.mxColBin.LinkPressedAfter += new SAPbouiCOM._IColumnEvents_LinkPressedAfterEventHandler(this.mxColBin_LinkPressedAfter);
            this.mxColBin.LinkPressedBefore += new SAPbouiCOM._IColumnEvents_LinkPressedBeforeEventHandler(this.mxColBin_LinkPressedBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.UnloadAfter += new SAPbouiCOM.Framework.FormBase.UnloadAfterHandler(this.Form_UnloadAfter);
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddBefore += new DataAddBeforeHandler(this.Form_DataAddBefore);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_Resize);
        }

        private void OnCustomInitialize()
        {

            this.btnRefresh.Item.Top = this.btnCancel.Item.Top;
            this.btnRefresh.Item.Height = this.btnCancel.Item.Height;
            this.btnRefresh.Item.Width = this.btnCancel.Item.Width;
            this.btnRefresh.Item.Left = this.btnCancel.Item.Left + this.btnCancel.Item.Width + 10;
        }

        private void Form_Resize(SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private void Form_UnloadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                BINLocation = string.Empty;
                mxRow = string.Empty;
            }
            catch (Exception ex)
            { throw ex; }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                for (int i = 1; i < mxContents.RowCount; i++)
                {
                    mxAxcBin = (SAPbouiCOM.EditText)mxAxcColBin.Cells.Item(i).Specific;
                    mxAxcBin.Value = "";
                }
            }
            catch (Exception)
            { //throw ex; 
            }

        }


        private void Form_DataAddBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if ((btTrans.Caption == "N" || btTrans.Caption == "V" || btTrans.Caption == "C") && !Application.SBO_Application.Forms.ActiveForm.Title.Contains("Cancellation"))
                {
                    BINLocation = string.Empty;
                    mxRow = string.Empty;
                    

                    //Application.SBO_Application.Forms.ActiveForm.Freeze(true);
                    for (int i = 1; i < mxContents.RowCount; i++)
                    {
                        Series = this.cbSeries.Selected.Value;
                        DocNum = this.edDocNum.Value;

                        //mxColBin.Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Linked, 0);          
                        BIN_Location_Auto_Select_Addon.Forms.frmBINAllocationIssue.Series = Series;

                        BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "SELECT * FROM \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                          "WHERE \"CompanyDB\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "' " + Environment.NewLine +
                          "AND \"UserSign\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "' " + Environment.NewLine +
                          "AND \"Series\" = '" + Series + "' " + Environment.NewLine +
                          "AND \"DocNum\" = '" + DocNum + "' " + Environment.NewLine +
                          "AND \"ObjType\" = '15'";
                        BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);

                        if (BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.RecordCount > 0)
                        {
                            while (!BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.EoF)
                            {
                                mxRow = BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.Fields.Item("RowNum").Value.ToString();
                                BINLocation = BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.Fields.Item("BINCode").Value.ToString();

                                if (!string.IsNullOrEmpty(BINLocation) && !string.IsNullOrEmpty(mxRow))
                                {
                                    ((SAPbouiCOM.EditText)(mxContents.GetCellSpecific("U_AXC_BINLocation", Convert.ToInt16(mxRow)))).Value = BINLocation; 
                                }

                                BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.MoveNext();
                            }
                        }
                    }
                    //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                throw ex;
            }
        }

        private void btnRefreshClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal) //commented
        {
            try
            {
                if ((btTrans.Caption == "N" || btTrans.Caption == "V" || btTrans.Caption == "C") && !Application.SBO_Application.Forms.ActiveForm.Title.Contains("Cancellation"))
                {
                    BINLocation = string.Empty;
                    mxRow = string.Empty;


                    //Application.SBO_Application.Forms.ActiveForm.Freeze(true);
                    for (int i = 1; i < mxContents.RowCount; i++)
                    {
                        Series = this.cbSeries.Selected.Value;
                        DocNum = this.edDocNum.Value;

                        //mxColBin.Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Linked, 0);
                        BIN_Location_Auto_Select_Addon.Forms.frmBINAllocationIssue.Series = Series;

                        BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "SELECT * FROM \"AXXIS_SAPB1\".\"AXXIS_TB_BINAutoSelect\" " + Environment.NewLine +
                          "WHERE \"CompanyDB\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.CompanyDB + "' " + Environment.NewLine +
                          "AND \"UserSign\" = '" + BINLocationAutoSelect.Helpers.GlobalVar.myCompany.UserSignature + "' " + Environment.NewLine +
                          "AND \"Series\" = '" + Series + "' " + Environment.NewLine +
                          "AND \"DocNum\" = '" + DocNum + "' " + Environment.NewLine +
                          "AND \"ObjType\" = '15'";
                        BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);

                        if (BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.RecordCount > 0)
                        {
                            while (!BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.EoF)
                            {
                                mxRow = BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.Fields.Item("RowNum").Value.ToString();
                                BINLocation = BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.Fields.Item("BINCode").Value.ToString();

                                if (!string.IsNullOrEmpty(BINLocation) && !string.IsNullOrEmpty(mxRow))
                                {
                                    ((SAPbouiCOM.EditText)(mxContents.GetCellSpecific("U_AXC_BINLocation", Convert.ToInt16(mxRow)))).Value = BINLocation;
                                }

                                BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.MoveNext();
                            }
                        }
                        //var binbtTrans = ((SAPbouiCOM.Button)(this.GetItem("1470000001").Specific));
                        
                    }
                    //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
                throw ex;
            }
            
        }
        private void btTrans_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        private void btTrans_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal) //commented
        {
            try
            {

            }
            catch (Exception)
            {
                //Application.SBO_Application.MessageBox(ex.ToString());

            }
        }
        private void btTrans_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (incErr > 0)
                {
                    BubbleEvent = false;
                }
            }
            catch (Exception)
            {
                BubbleEvent = false;
                //throw ex;                
            }
        }

        private void mxContents_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "U_AXC_BINLocation")
                {
                    if (oldbincod != ((SAPbouiCOM.EditText)this.mxAxcColBin.Cells.Item(pVal.Row).Specific).Value.ToString())
                    {
                        Success = false;
                        msgProcess = "Reprocessing Bin Allocation, please wait...";
                    }
                }
            }
            catch (Exception)
            {
                //Application.SBO_Application.MessageBox(ex.ToString());
            }
        }
        private void mxColBin_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                BINLocationAutoSelect.Program.SelectedSeries = this.cbSeries.Selected.Value;
            }
            catch (Exception ex)
            {

            }
        }

        private void mxColBin_LinkPressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //Application.SBO_Application.Forms.ActiveForm.Freeze(true);

                if (!String.IsNullOrEmpty(bincod) && !BinNotExist && BINLocationAutoSelect.Helpers.GlobalVar._frmAutoPop)
                {
                    BINLocationAutoSelect.Helpers.GlobalVar.PopBinReceipt(frmType, frmCnt, bincod, binqty);
                }
                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);                 
            }
            catch (Exception)
            {
                //Application.SBO_Application.MessageBox(ex.ToString());
                //Application.SBO_Application.Forms.ActiveForm.Freeze(false);
            }
        }
        private void mxContents_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "U_AXC_BINLocation")
                {
                    oldbincod = ((SAPbouiCOM.EditText)this.mxAxcColBin.Cells.Item(pVal.Row).Specific).Value.ToString();
                }
            }
            catch (Exception)
            {
                //Application.SBO_Application.MessageBox(ex.ToString());
            }
        }
        //private void linkbtBinCol_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal) //commented
        //{
        //    try
        //    {

        //    }
        //    catch (Exception)
        //    {
        //        //Application.SBO_Application.MessageBox(ex.ToString());

        //    }
        //}
    }
}
