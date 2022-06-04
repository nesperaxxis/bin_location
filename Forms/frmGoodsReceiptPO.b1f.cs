using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Drawing;

namespace BIN_Location_Auto_Select_Addon.Forms
{
    [FormAttribute("143", "Forms/frmGoodsReceiptPO.b1f")]
    class frmGoodsReceiptPO : SystemFormBase
    {
        #region SAP Objects
        public SAPbouiCOM.ButtonCombo btTrans { get; set; }
        public SAPbouiCOM.ComboBox cbSeries { get; set; }
        public SAPbouiCOM.EditText edDocNum { get; set; }
        public SAPbouiCOM.Matrix mxContents { get; set; }
        public SAPbouiCOM.Column mxColItem { get; set; }
        public SAPbouiCOM.Column mxColQty { get; set; }
        public SAPbouiCOM.Column mxColBin { get; set; }
        public SAPbouiCOM.Column mxAxcColBin { get; set; }
        public SAPbouiCOM.LinkedButton mxBinLink { get; set; }
        public SAPbouiCOM.Form _frmGRPO { get; set; }
        
        #endregion SAP Objects

        #region Variables
        public static string Series { get; set; }
        public static string DocNum { get; set; }
        public static string BINLocation { get; set; }
        public static bool BinNotExist { get; set; }
        public static string mxRow { get; set; }
        public static int mxRowErr { get; set; }
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
        public frmGoodsReceiptPO()
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
            this.mxColQty = (SAPbouiCOM.Column)(this.mxContents.Columns.Item("1")); 
            this.mxColQty = (SAPbouiCOM.Column)(this.mxContents.Columns.Item("11")); 
            this.mxColBin = (SAPbouiCOM.Column)(this.mxContents.Columns.Item("1470002149")); //BinLoc
            this.mxAxcColBin = (SAPbouiCOM.Column)(this.mxContents.Columns.Item("U_AXC_BINLocation"));
            //this.mxBinLink = ((SAPbouiCOM.LinkedButton)(this.GetItem("1470002149").Specific));
            this.btTrans = ((SAPbouiCOM.ButtonCombo)(this.GetItem("1").Specific));
            this.btTrans.ClickBefore += new SAPbouiCOM._IButtonComboEvents_ClickBeforeEventHandler(this.btTrans_ClickBefore);
            this.btTrans.ClickAfter += new SAPbouiCOM._IButtonComboEvents_ClickAfterEventHandler(this.btTrans_ClickAfter);
            this.btTrans.PressedBefore += new SAPbouiCOM._IButtonComboEvents_PressedBeforeEventHandler(this.btTrans_PressedBefore);
            //this.btTrans.PressedAfter += new SAPbouiCOM._IButtonComboEvents_PressedAfterEventHandler(this.btTrans_PressedAfter);
            this.mxContents.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.mxContents_GotFocusAfter);
            this.mxContents.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.mxContents_LostFocusAfter);
            this.mxColBin.LinkPressedAfter += new SAPbouiCOM._IColumnEvents_LinkPressedAfterEventHandler(this.mxColBin_LinkPressedAfter);
            this.OnCustomInitialize();
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>

        public override void OnInitializeFormEvents()
        {
            //this.UnloadAfter += new SAPbouiCOM.Framework.FormBase.UnloadAfterHandler(this.Form_UnloadAfter);
            //this.DataAddBefore += new DataAddBeforeHandler(this.Form_DataAddBefore);
            //this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);            
        }

        private void OnCustomInitialize()
        {
            try
            {
                _frmGRPO = Application.SBO_Application.Forms.ActiveForm;
                frmCnt = _frmGRPO.TypeCount;
                frmType = _frmGRPO.TypeEx;
                BinNotExist = false;
                BINLocationAutoSelect.Helpers.GlobalVar._frmAutoPop = false;
                Success = false;
                msgProcess = "Processing Bin Allocation, please wait...";
            }
            catch(Exception)
            {
                //will not do anything
            }
 
        }
     
        private void mxContents_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "U_AXC_BINLocation")
                {
                    if( oldbincod != ((SAPbouiCOM.EditText)this.mxAxcColBin.Cells.Item(pVal.Row).Specific).Value.ToString())
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
        private void btTrans_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
                _frmGRPO = Application.SBO_Application.Forms.ActiveForm;
                if(!_frmGRPO.Title.Contains("Cancel"))
                {
                    if (!Success && (btTrans.Caption.Equals("C") || btTrans.Caption.Equals("N") || btTrans.Caption.Equals("V")))
                    {
                        if (mxContents.RowCount > 1)
                        {
                            string BinVal = String.Empty;
                            string BinQty = String.Empty;
                            int whiteBackColor = 0;
                            int redBackColor = 0;
                            incErr = 0;
                            incSuccess = 0;
                            for (int i = 1; i < mxContents.RowCount; i++)
                            {
                                BinVal = ((SAPbouiCOM.EditText)this.mxAxcColBin.Cells.Item(i).Specific).Value.ToString();
                                BinQty = ((SAPbouiCOM.EditText)this.mxColQty.Cells.Item(i).Specific).Value.ToString();
                                DocNum = edDocNum.Value.ToString();
                                Series = cbSeries.Value.ToString();
                                bincod = BinVal;
                                binqty = BinQty;

                                if (!String.IsNullOrEmpty(BinVal))
                                {
                                    //check if bincode exist

                                    BINLocationAutoSelect.Helpers.GlobalVar.oQuery = "SELECT * FROM \"OBIN\" " + Environment.NewLine +
                                      "WHERE \"BinCode\" = '" + BinVal + "' ";
                                    BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.DoQuery(BINLocationAutoSelect.Helpers.GlobalVar.oQuery);

                                    if (BINLocationAutoSelect.Helpers.GlobalVar.oRSQuery.RecordCount > 0)
                                    {
                                        Application.SBO_Application.SetStatusBarMessage(" Please wait while we process bin allocation...", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                        whiteBackColor = Color.White.R | (Color.White.G << 8) | (Color.White.B << 16);
                                        mxContents.CommonSetting.SetRowBackColor(i, whiteBackColor);
                                        BinNotExist = false;
                                        BINLocationAutoSelect.Helpers.GlobalVar.callBinReceipt(frmType, frmCnt, i);
                                        if (_frmGRPO.Selected && BINLocationAutoSelect.Helpers.GlobalVar._frmAutoPop)
                                        {
                                            mxColBin.Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Linked, 0);
                                        }
                                        incSuccess++;
                                    }
                                    else
                                    {

                                        redBackColor = Color.Red.R | (Color.Red.G << 8) | (Color.Red.B << 16);
                                        mxContents.CommonSetting.SetRowBackColor(i, redBackColor); //201205
                                        BinNotExist = true;
                                        mxRowErr = i;
                                        BubbleEvent = false;
                                        incErr++;
                                    }
                                }
                            }
                            if (incErr > 0 || incSuccess > 0)
                            {
                                if (incErr > 0)
                                {
                                    Application.SBO_Application.SetStatusBarMessage(" Some Bin Locations does not exist. ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    BubbleEvent = false;
                                    Success = false;
                                    msgProcess = "Reprocessing Bin Allocation, please wait...";
                                }
                                else
                                {
                                    Application.SBO_Application.SetStatusBarMessage(" Process Bin alloocations successful. ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    Success = true;
                                    btTrans.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                            }
                        } //Row Count End
                    }
                }
            }
            catch (Exception)
            {
               // Application.SBO_Application.MessageBox(ex.ToString());
                BubbleEvent = false;
            }
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
                if(incErr > 0)
                {
                    BubbleEvent = false;
                }
            }
            catch(Exception)
            {
                BubbleEvent = false;
                //throw ex;                
            }
        }
    }
}
