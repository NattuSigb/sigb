using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using General;
using General.Classes;
//using DepartmentWiseWharehouseManagement;
using SIGBNat;
using General.Extensions;
using log4net;
using log4net.Config;
using log4net.Repository.Hierarchy;
using SAPbouiCOM;
using System.Reflection;
using static General.Classes.@enum;


namespace SIGBNat
{
    class SAPMain : Connection
    {
        public static ILog logger;
        SAPbouiCOM.Form oForm;

        #region Variables
        //the below lines will be given then only on the PreparMenus we can access objclsComman sowhen copy&paste 
        //it will show you the error so copying the below lines from the Reference code
        clsCommon objclsComman = new clsCommon();

        BPIMaster objBPIMaster = new BPIMaster();
        CCBI objCCBI = new CCBI();
        IMD objIMD = new IMD();
        //clsDWHM objclsDWHM = new clsDWHM();
        //clsPurchase objclsPurchase = new clsPurchase();
        //clsSalesorder objclsSales = new clsSalesorder();

        #endregion

        #region ConstruBPNamer

        public SAPMain()
        {
            InitLogger();
            ConnectToSAPApplication();
            PrepareMenus();
            PrepareEvents();
        }

        private void PrepareMenus()
        {
            logger.DebugFormat("> {0}", nameof(PrepareMenus));
            string superUser = objclsComman.SelectRecord("SELECT SuperUser FROM OUSR WHERE UserId='" + oCompany.UserSignature + "'");
            if (superUser == YesNoEnum.Y.ToString())
            {
                objclsComman.AddMenu(BoMenuType.mt_STRING, Convert.ToString((int)SAPMenuEnum.SystemInitialisation), SAPCustomFormUIDEnum.BPIUDO.ToString(), SAPCustomFormUIDEnum.BPIUDO.ToDescription(), "BPIUDO", 0);
            }
            if (isLiceseExpired == false)
            {
                objclsComman.AddMenu(BoMenuType.mt_POPUP, Convert.ToString((int)SAPMenuEnum.Modules), SAPCustomFormUIDEnum.BPIADO.ToString(), SAPCustomFormUIDEnum.BPIADO.ToDescription(), "BPIUDO", 16);//SAPCustomFormUIDEnum.BPIADO.ToString()
                objclsComman.AddMenu(BoMenuType.mt_STRING, SAPCustomFormUIDEnum.BPIADO.ToString(), SAPCustomFormUIDEnum.BPIM.ToString(), SAPCustomFormUIDEnum.BPIM.ToDescription(), "BPIUDO", 0);
                objclsComman.AddMenu(BoMenuType.mt_STRING, SAPCustomFormUIDEnum.BPIADO.ToString(), SAPCustomFormUIDEnum.CCBI.ToString(), SAPCustomFormUIDEnum.CCBI.ToDescription(), "BPIUDO", 1);
                
            }
        }

        void RemoveMenusFromSAP()
        {
            objclsComman.RemoveMenu(SAPCustomFormUIDEnum.BPIUDO.ToString());
            objclsComman.RemoveMenu(SAPCustomFormUIDEnum.BPIADO.ToString());
            objclsComman.RemoveMenu(SAPCustomFormUIDEnum.BPIM.ToString());
            

        }

        /// <summary>
        /// Create Event Handler 
        /// </summary>
        void PrepareEvents()
        {
            logger.DebugFormat("> {0} ", nameof(PrepareEvents));
           oApplication.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(oApplication_ItemEvent);
            oApplication.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(oApplication_MenuEvent);
            //oApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(oApplication_FormDataEvent);
            //oApplication.AppEvent += new _IApplicationEvents_AppEventEventHandler(oApplication_AppEvent);
            oApplication.LayoutKeyEvent += new _IApplicationEvents_LayoutKeyEventEventHandler(oApplication_LayoutKeyEvent);

        }

        #endregion

        #region Events

   
        void oApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (pVal.FormTypeEx == SAPCustomFormUIDEnum.BPIM.ToString())
                {
                    objBPIMaster.ItemEvent(ref pVal, out BubbleEvent);
                }
                 if (pVal.FormTypeEx == SAPCustomFormUIDEnum.CCBI.ToString())
                {
                    objCCBI.ItemEvent(ref pVal, out BubbleEvent);
                }
                 if (pVal.FormTypeEx == Convert.ToString((int)SAPFormUIDEnum.ItemMasterData))
                {
                    objIMD.ItemEvent(ref pVal, out BubbleEvent);
                }
                //else if (pVal.FormTypeEx == SAPCustomFormUIDEnum.SOF.ToString() || (oForm != null && oForm.TypeEx == SAPCustomFormUIDEnum.SOF.ToString()))

                //{
                //    objCCBI.ItemEvent(ref pVal, out BubbleEvent);
                //}//else if (pVal.FormTypeEx == Convert.ToString((int)SAPFormUIDEnum.PurchaseOrder))
                //{
                //    objclsPurchase.ItemEvent(ref pVal, out BubbleEvent);
                //}
                //else if (pVal.FormTypeEx == Convert.ToString((int)SAPFormUIDEnum.SalesOrder))
                //{
                //    objclsSales.ItemEvent(ref pVal, out BubbleEvent);
                //}

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        //.....
        //.....

        void oApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                try
                {
                    oForm = oApplication.Forms.ActiveForm;
                }
                catch


                {
                    oForm = null;
                }

                if (pVal.BeforeAction == true)
                {
                    if (pVal.MenuUID == SAPCustomFormUIDEnum.BPIUDO.ToString())
                    {
                        objclsComman.CreateDataBase();
                        oApplication.StatusBar.SetText("Object Tables Created successfully!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                }

                //Below is the code for accessing Menu Events

                if (pVal.MenuUID == SAPCustomFormUIDEnum.BPIM.ToString() || (oForm != null && oForm.TypeEx == SAPCustomFormUIDEnum.BPIM.ToString()))
                {
                    //BPIMaster.(ref pVal, out BubbleEvent);
                     objBPIMaster.MenuEvent(ref pVal, out BubbleEvent);
                }

                 if (pVal.MenuUID == SAPCustomFormUIDEnum.CCBI.ToString() || (oForm != null && oForm.TypeEx == SAPCustomFormUIDEnum.CCBI.ToString()))
                {
                    //BPIMaster.(ref pVal, out BubbleEvent);
                    objCCBI.MenuEvent(ref pVal, out BubbleEvent);
                }

                if (pVal.MenuUID == Convert.ToString((int)SAPFormUIDEnum.ItemMasterData) || (oForm != null && oForm.TypeEx == Convert.ToString((int)SAPFormUIDEnum.ItemMasterData)))
                {
                    //BPIMaster.(ref pVal, out BubbleEvent);
                    objIMD.MenuEvent(ref pVal, out BubbleEvent);
                }


            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /*
        void oApplication_FormDataEvent( ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.FormTypeEx == SAPCustomFormUIDEnum.DWHM.ToString())
                {
                    objclsDWHM.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }
                else if (BusinessObjectInfo.FormTypeEx == Convert.ToString((int)SAPFormUIDEnum.PurchaseOrder))
                {
                    objclsPurchase.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }
                //else if(BusinessObjectInfo.FormTypeEx == Convert.ToString((int)SAPCustomFormUIDEnum.SalesOrder))
                //{
                //    objclsSales.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                //}
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }     
        void oApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {


                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    oApplication.SetStatusBarMessage(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + " Addon... Shutdown Event has been caught" + Environment.NewLine + "Terminating Add On...", BoMessageTime.bmt_Short, false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);

                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }
*/
        void oApplication_LayoutKeyEvent(ref LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (eventInfo.BeforeAction == true)
            {
                if (eventInfo.FormUID.Contains(SAPCustomFormUIDEnum.BPIM.ToString())
                   )
                {
                    eventInfo.LayoutKey = clsVariables.DocEntry;
                }


            }
        }

        #endregion

        #region Methods

        /// <summary>
        ///     Configure log4net system based on application configuration setting
        /// </summary>
        private static void InitLogger()
        {
            XmlConfigurator.Configure();
            ((Hierarchy)LogManager.GetRepository()).RaiseConfigurationChanged(EventArgs.Empty);
            logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            logger.Info("Logger initialzed.");
        }

        #endregion
    }
}