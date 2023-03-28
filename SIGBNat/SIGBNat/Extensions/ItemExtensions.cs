using System.Runtime.InteropServices;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using General.Classes;
using static General.Classes.@enum;

namespace General.Extensions
{
    public static class ItemExtensions
    {
        /// <summary>
        /// Enable control
        /// </summary>
        /// <param name="item">item control</param>
        public static void Enable(this Item item)
        {
            item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPCommonMaskModeEnum.All, BoModeVisualBehavior.mvb_True);
        }

        public static void EnableinFindMode(this Item item)
        {
            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
        }

        /// <summary>
        /// Enable control in Add Mode Only
        /// </summary>
        /// <param name="item">item control</param>
        public static void EnableinAddMode(this Item item)
        {
            item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPCommonMaskModeEnum.Ok, BoModeVisualBehavior.mvb_False);
        }

        /// <summary>
        /// Disable control
        /// </summary>
        /// <param name="item">item control</param>
        public static void Disable(this Item item)
        {
            item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)SAPCommonMaskModeEnum.All, BoModeVisualBehavior.mvb_False);
        }


    }
}