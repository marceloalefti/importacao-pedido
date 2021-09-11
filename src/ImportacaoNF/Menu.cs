using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SLT.ImportacaoNF
{
    class Menu
    {
        private SAPbouiCOM.Application SAPApp;
        private const string MenuBaseUID = "Alefti";
        private const string MenuImportacaoUID = "SLT.ImportacaoNF.frmImportacao";

        public void AddMenuItems()
        {
            SAPApp = SAPbouiCOM.Framework.Application.SBO_Application;

            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            oMenus = SAPApp.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;

            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SAPApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = SAPApp.Menus.Item("43520"); // modules'

            //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = MenuBaseUID;
            oCreationPackage.String = "Alefti Tools";
            oCreationPackage.Enabled = true;
            oCreationPackage.Image = Path.Combine(System.Windows.Forms.Application.StartupPath, @"assets\cogs-icon-16.png");
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;
            AddMenu(oMenus, oCreationPackage);

            oMenuItem = SAPApp.Menus.Item(MenuBaseUID);
            oMenus = oMenuItem.SubMenus;

            // Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = MenuImportacaoUID;
            oCreationPackage.String = "Controle de importação";
            AddMenu(oMenus, oCreationPackage);
        }

        public void RemoveMenuItems()
        {
            SAPbouiCOM.Menus oMenus = (SAPbouiCOM.Menus)SAPApp.Menus;
            RemoveMenu(oMenus, MenuImportacaoUID);
            RemoveMenu(oMenus, MenuBaseUID);
        }

        private void AddMenu(SAPbouiCOM.Menus oMenus, SAPbouiCOM.MenuCreationParams oCreationPackage)
        {
            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                SAPApp.SetStatusBarMessage("O menu já existe!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void RemoveMenu(SAPbouiCOM.Menus oMenus, string menuUID)
        {
            try
            {
                //  If the manu already exists this code will fail
                oMenus.RemoveEx(menuUID);
            }
            catch (Exception ex)
            {
                SAPApp.SetStatusBarMessage("Não foi possível remover o menu!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "SLT.ImportacaoNF.frmImportacao")
                {
                    var form = new frmImportacao();
                    //form.Show();
                }
            }
            catch (Exception ex)
            {
                SAPApp.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
