using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using General.Classes;
using SAPbouiCOM;
using SAPbobsCOM;
using SIGBNat;
using static General.Classes.@enum;
using General.Extensions;

namespace SIGBNat
{
    class IMD : Connection
    {
        public const string headerTable = "OITM";
        const string formMenuUID = "OITM";
        SAPbouiCOM.Form oForm;
        clsCommon objclsCommon = new clsCommon();
        SAPbouiCOM.Item oNewItem;
        SAPbouiCOM.Item oItem;
        SAPbouiCOM.Button oButton;
        public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region BeforeAction == true

                if (pVal.BeforeAction == true)
                {
                    try
                    {
                        if (pVal.MenuUID == formMenuUID || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRecord))
                        {
                            oForm = oApplication.Forms.ActiveForm;
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                //Record is directly added without validation
                                BubbleEvent = false;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = true : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

                #region BeforeAction == false

                if (pVal.BeforeAction == false)
                {
                    try
                    {
                        if (pVal.MenuUID == formMenuUID)
                        {
                            LoadForm(pVal.MenuUID);
                        }




                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        public void ItemEvent(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                #region Before_Action == true

                if (pVal.Before_Action == true)
                {

                }
                #endregion

             
                if (pVal.Before_Action == false)
                {

                     

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {

                        oForm = oApplication.Forms.Item(pVal.FormUID);


                        oNewItem = oForm.Items.Add("ExportBtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        oItem = oForm.Items.Item("2");
                        oNewItem.Left = oItem.Left + 180;
                        oNewItem.Width = oItem.Width;
                        oNewItem.Top = oItem.Top;
                        oNewItem.Height = oItem.Height;
                        oNewItem.Enabled = true;
                        oButton = oNewItem.Specific;
                        oButton.Caption = "Link to Purchase Order";
                    }
                }

            }
            catch (Exception e)
            { 
            }


        }


        #region Methods
        public void LoadForm(string MenuID)
        {
            clsVariables.boolCFLSelected = false;

            if (MenuID == formMenuUID)
            {
                string formUID = "";
                objclsCommon.LoadXML(MenuID, "", string.Empty, SAPbouiCOM.BoFormMode.fm_ADD_MODE);
                oForm = oApplication.Forms.ActiveForm;
                oForm.DataSources.UserDataSources.Add("Close", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Item("Close").Value = "N";

            }

            #endregion
        }
    }
}
    
