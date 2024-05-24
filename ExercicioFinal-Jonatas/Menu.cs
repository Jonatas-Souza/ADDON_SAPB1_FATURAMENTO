using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExercicioFinal_Jonatas
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "ExercicioFinal_Jonatas";
            oCreationPackage.String = "Jonatas - Exercício Final";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;
            oCreationPackage.Image = Environment.CurrentDirectory + @"\addon_Icon.png";

            oMenus = oMenuItem.SubMenus;

            try
            {
                if (oMenus.Exists(oCreationPackage.UniqueID))
                {
                    oMenus.RemoveEx(oCreationPackage.UniqueID);
                }
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("ExercicioFinal_Jonatas");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ExercicioFinal_Jonatas.Form1";
                oCreationPackage.String = "Assistente de geração de notas fiscais";

                if (oMenus.Exists(oCreationPackage.UniqueID))
                {
                    oMenus.RemoveEx(oCreationPackage.UniqueID);
                }

                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "ExercicioFinal_Jonatas.Form1")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
