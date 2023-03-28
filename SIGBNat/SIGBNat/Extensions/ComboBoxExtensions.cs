using System.Runtime.InteropServices;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
//using DepartmentWiseWharehouseManagement;
using SIGBNat;

namespace General.Extensions
{
    public static class ComboBoxExtensions
    {
        /// <summary>
        ///     Add valid values into comboBox
        /// </summary>
        /// <param name="comboBox">the combobox to add valid values</param>
        /// <param name="recordsetQuery">query string to query valid values</param>
        public static void AddValidQuery(this ComboBox comboBox, SAPbobsCOM.Recordset recordset)
        {
            try
            {

                while (!recordset.EoF)
                {
                    try
                    {
                        comboBox.ValidValues.Add(recordset.Fields.Item(0).Value.ToString(),
                        recordset.Fields.Item(1).Value.ToString());
                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.WarnFormat("{0}: {1} {2}", ex.Message, recordset.Fields.Item(0).Value.ToString(), recordset.Fields.Item(1).Value.ToString());
                    }
                    recordset.MoveNext();
                }
            }
            catch (Exception e)
            {
                SAPMain.logger.Error(e.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        ///     Extension method to clear valid values for a comboBox
        /// </summary>
        /// <param name="comboBox">the comboBox to clear valid values</param>
        public static void ClearValidValues(this ComboBox comboBox)
        {
            for (int i = comboBox.ValidValues.Count; i >= 1; i--)
            {
                comboBox.ValidValues.Remove(i - 1, BoSearchKey.psk_Index);
            }
        }
    }
}