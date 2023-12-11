using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_JE.Common
{
    class clsRightClickEvent
    {
        SAPbouiCOM.Form objform;
        clsGlobalMethods objglobalMethods= new clsGlobalMethods();
        SAPbouiCOM.ComboBox ocombo;
        SAPbouiCOM.Matrix objmatrix;
        string strsql;
        SAPbobsCOM.Recordset objrs;

        public void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "133":
                    case "REVREC":
                        {                            
                            //RevenueRecognition_RightClickEvent(ref eventInfo,ref BubbleEvent);
                            break;
                        }
                   
                }
            }
            catch (Exception ex)
            {
            }
        }


        private void RightClickMenu_Add(string MainMenu, string NewMenuID, string NewMenuName, int position)
        {
            SAPbouiCOM.Menus omenus;
            SAPbouiCOM.MenuItem omenuitem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;
            oCreationPackage =(SAPbouiCOM.MenuCreationParams)clsModule.objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            omenuitem = clsModule.objaddon.objapplication.Menus.Item(MainMenu); // Data'
            if (!omenuitem.SubMenus.Exists(NewMenuID))
            {
                oCreationPackage.UniqueID = NewMenuID;
                oCreationPackage.String = NewMenuName;
                oCreationPackage.Position = position;
                oCreationPackage.Enabled = true;
                omenus = omenuitem.SubMenus;
                omenus.AddEx(oCreationPackage);
            }
        }

        private void RightClickMenu_Delete(string MainMenu, string NewMenuID)
        {
            SAPbouiCOM.MenuItem omenuitem;
            omenuitem = clsModule.objaddon.objapplication.Menus.Item(MainMenu); // Data'
            if (omenuitem.SubMenus.Exists(NewMenuID))
            {
                clsModule.objaddon.objapplication.Menus.RemoveEx(NewMenuID);
            }
        }

        private void RevenueRecognition_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Matrix Matrix0;
                objform =clsModule. objaddon.objapplication.Forms.ActiveForm;
                Matrix0 =(SAPbouiCOM.Matrix) objform.Items.Item("mtxcont").Specific;
                if (eventInfo.BeforeAction)
                {
                    objform.EnableMenu("1283", false);
                    objform.EnableMenu("1285", false);
                    if (eventInfo.ItemUID == "")
                    {
                        if (((SAPbouiCOM.EditText)objform.Items.Item("ttranid").Specific).String != "" & ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String=="")                        
                            objform.EnableMenu("1284", true);
                        else objform.EnableMenu("1284", false);
                    }
                    objform.EnableMenu("1286", false);
                    
                    try
                    {
                          // Copy Table
                        if (Matrix0.Item.Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                        {
                            if (eventInfo.ItemUID== "mtxcont") objform.EnableMenu("784", true); //Copy Table
                            if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific).String != "")
                            {
                                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) objform.EnableMenu("1293", true);
                                else objform.EnableMenu("1293", false); // Remove Row Menu
                                objform.EnableMenu("772", true);  // Copy                               
                            }
                            else
                            {
                                objform.EnableMenu("772", false);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (((SAPbouiCOM.EditText)objform.Items.Item(eventInfo.ItemUID).Specific).String != "")
                        {
                            objform.EnableMenu("772", true);  // Copy
                        }
                        else
                        {
                            objform.EnableMenu("772", false);
                        }
                    }
                }
                else
                {                    
                    if (((SAPbouiCOM.EditText)objform.Items.Item("ttranid").Specific).String != "")
                    {
                         objform.EnableMenu("1293", false); // Remove Row Menu
                        if (eventInfo.ItemUID=="")objform.EnableMenu("1284", true);
                    }
                    else
                    {
                        
                    }
                    objform.EnableMenu("784", false);
                    objform.EnableMenu("1284", false);
                    objform.EnableMenu("772", false);
                    objform.EnableMenu("1293", false);
                }
                
            }
            catch (Exception ex)
            {
            }
        }

        private void ProjectMaster_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Matrix Matrix0;
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mtxcont").Specific;
                if (eventInfo.BeforeAction)
                {
                    if (objform.Mode== SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.EnableMenu("1283", true); // Remove
                    else objform.EnableMenu("1283", false);
                    objform.EnableMenu("1285", false);
                    objform.EnableMenu("1284", false);
                    objform.EnableMenu("1286", false);
                    if (eventInfo.ColUID == "#")
                    {
                        objform.EnableMenu("1293", true); // Remove Row Menu
                    }
                    try
                    {
                        // Copy Table
                        if (Matrix0.Item.Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                        {
                            if (eventInfo.ItemUID == "mtxcont") objform.EnableMenu("784", true); //Copy Table
                            if (((SAPbouiCOM.EditText)Matrix0.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific).String != "")
                            {
                                objform.EnableMenu("772", true);  // Copy                               
                            }
                            else
                            {
                                objform.EnableMenu("772", false);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (((SAPbouiCOM.EditText)objform.Items.Item(eventInfo.ItemUID).Specific).String != "")
                        {
                            objform.EnableMenu("772", true);  // Copy
                        }
                        else
                        {
                            objform.EnableMenu("772", false);
                        }
                    }
                }
                else
                {
                    objform.EnableMenu("1293", false); // Remove Row Menu

                    objform.EnableMenu("784", false);
                    objform.EnableMenu("1284", false);
                    objform.EnableMenu("772", false);
                    objform.EnableMenu("1293", false);
                }

            }
            catch (Exception ex)
            {
            }
        }

    }
}
