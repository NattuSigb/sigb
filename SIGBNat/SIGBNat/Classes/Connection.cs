using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using DepartmentWiseWharehouseManagement;
//using DepartmentWiseWharehouseManagement.Classes;
using SIGBNat;

namespace General.Classes
{
    class Connection
    {
        #region SAP Objects

        public static SAPbouiCOM.Application oApplication;
        public static SAPbobsCOM.Company oCompany;
        public static bool isLiceseExpired;
        //public static string salt;

        #endregion

        #region Methods

        /// <summary>
        /// Connect to SAP application
        /// </summary>
        public void ConnectToSAPApplication()
        {
            try
            {
               // SAPMain.logger.DebugFormat("> {0}", nameof(ConnectToSAPApplication));
                string devConnectionString = string.Empty;
                try
                {
                    devConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                }
                catch
                {
                    devConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                }
                var sboGuiApi = new SAPbouiCOM.SboGuiApi();
                sboGuiApi.Connect(devConnectionString);

                oApplication = sboGuiApi.GetApplication();
                oCompany = (SAPbobsCOM.Company)oApplication.Company.GetDICompany();
                if (isLiceseExpired == false)
                {
                    oApplication.StatusBar.SetText(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + " addon connect successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Connect Application Error: " + ex.Message);
            }
        }

        public string SelectRecord(string query)
        {
            SAPbobsCOM.Recordset oRs = null;
            string value = "";

            try
            {
                oRs = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRs.DoQuery(query);

                if (oRs.EoF == false)
                {
                    value = oRs.Fields.Item(0).Value.ToString();
                }
                return value;
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return value;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                oRs = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }



        #endregion
    }
}





       



