using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Automatic_JE.Business_Objects;
using SAPbouiCOM.Framework;
using System.Diagnostics;

namespace Automatic_JE.Common
{
    class clsAddon
    {
        public SAPbouiCOM.Application objapplication;
        public SAPbobsCOM.Company objcompany;
        public clsMenuEvent objmenuevent;
        public clsRightClickEvent objrightclickevent;
        public clsGlobalMethods objglobalmethods;     
        private SAPbouiCOM.Form objform,udfForm;
        public string addonName = "Service JE";
        string strsql= "";
        private SAPbobsCOM.Recordset objrs;
        bool print_close = false;
        public clsDocumentSettings documentSettings;
        public clsGRPO gRPO;
        public clsAPInvoice aPInvoice;
        //public bool HANA = true;
        public bool HANA = false;

        public string[] HWKEY   =  { "L1653539483", "X1211807750", "E0154677852" };

        #region Constructor
        public clsAddon()
        {
            
        }
        #endregion

        public void Intialize(string[] args)
        {
            try
            {
                Application oapplication;
                if ((args.Length < 1))
                    oapplication = new Application();
                else
                    oapplication = new Application(args[0]);
               
                objapplication = Application.SBO_Application;
                if (isValidLicense())
                {
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    objcompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                    Create_DatabaseFields(); // UDF & UDO Creation Part    
                    Menu(); // Menu Creation Part
                    Create_Objects(); // Object Creation Part
                    Add_Authorizations(); //User Permissions

                    objapplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(objapplication_AppEvent);
                    objapplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(objapplication_MenuEvent);
                    objapplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objapplication_ItemEvent);
                    objapplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(ref FormDataEvent);
                    //objapplication.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(objapplication_ProgressBarEvent);
                    //objapplication.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(objapplication_StatusBarEvent);
                    objapplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(objapplication_RightClickEvent);

                    objapplication.StatusBar.SetText(addonName +" Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    oapplication.Run();
                }
                else
                {
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                    //throw new Exception(objcompany.GetLastErrorDescription());
                }
            }
            // System.Windows.Forms.Application.Run()
            catch (Exception ex)
            {
                objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }      
        
        public bool isValidLicense()
        {
            try
            {
                if (HANA)
                {
                    try
                    {
                        if (objapplication.Forms.ActiveForm.TypeCount > 0)
                        {
                            for (int i = 0; i <= objapplication.Forms.ActiveForm.TypeCount - 1; i++)
                                objapplication.Forms.ActiveForm.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                // If Not HANA Then
                // objapplication.Menus.Item("1030").Activate()
                // End If
                objapplication.Menus.Item("257").Activate();
                SAPbouiCOM.EditText objedit= (SAPbouiCOM.EditText)objapplication.Forms.ActiveForm.Items.Item("79").Specific;
              
                string CrrHWKEY = objedit.Value.ToString();
                objapplication.Forms.ActiveForm.Close();

                for (int i = 0; i <= HWKEY.Length - 1; i++)
                {
                    //string key = HWKEY[i];
                    if (HWKEY[i] == CrrHWKEY)
                    {
                        return true;
                    }
                        
                }
                
                System.Windows.Forms.MessageBox.Show("Installing Add-On failed due to License mismatch");
                //objapplication.MessageBox("Installing Add-On failed due to License mismatch", 1, "Ok", "", "");
                //Interaction.MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management");

                return false;
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            return true;
        }

        public void Create_Objects()
        {
            objmenuevent = new clsMenuEvent();
            objrightclickevent = new clsRightClickEvent();
            objglobalmethods = new clsGlobalMethods();
            documentSettings = new clsDocumentSettings();
            gRPO = new clsGRPO();
            aPInvoice = new clsAPInvoice();

        }

        private void Create_DatabaseFields()
        {
            // If objapplication.Company.UserName.ToString.ToUpper <> "MANAGER" Then

            // If objapplication.MessageBox("Do you want to execute the field Creations?", 2, "Yes", "No") <> 1 Then Exit Sub
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            var objtable = new clsTable();
            objtable.FieldCreation();
            // End If

        }

        public void Add_Authorizations()
        {
            try
            {
                //clsModule.objaddon.objglobalmethods.AddToPermissionTree("Altrocks Tech", "ATPL_ADD-ON", "", "", 'Y');
                //clsModule.objaddon.objglobalmethods.AddToPermissionTree("Revenue Recognition", "ATPL_REVREG", "", "ATPL_ADD-ON", 'Y');
                //clsModule.objaddon.objglobalmethods.AddToPermissionTree("Project Master", "ATPL_PRJMSTR", "REVPRJMSTR", "ATPL_REVREG", 'Y');

            }
            catch (Exception ex)
            {

            }
        }

        #region Menu Creation Details

        private void Menu()
        {
            int Menucount = 1;
            //if (objapplication.Menus.Item("2048").SubMenus.Exists("REVRECO") & objapplication.Menus.Item("43525").SubMenus.Exists("REVPRJMSTR"))
            //    return;
            //Menucount = 1;// objapplication.Menus.Item("8448").SubMenus.Count;
            //CreateMenu("", Menucount, "Revenue Recognition", SAPbouiCOM.BoMenuType.mt_STRING, "REVRECO", "2048");

            //CreateMenu("", Menucount, "Revenue Master", SAPbouiCOM.BoMenuType.mt_POPUP, "REVRCG", "43525");  //Administration Module-> Setup
            //Menucount = 1;//Menu Inside   
            //CreateMenu("", Menucount, "Project Master", SAPbouiCOM.BoMenuType.mt_STRING, "REVPRJMSTR", "REVRCG"); Menucount += 1;

        }

        private void CreateMenu(string ImagePath, int Position, string DisplayName, SAPbouiCOM.BoMenuType MenuType, string UniqueID, string ParentMenuID)
        {
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuPackage;
                SAPbouiCOM.MenuItem parentmenu;
                parentmenu = objapplication.Menus.Item(ParentMenuID);
                if (parentmenu.SubMenus.Exists(UniqueID.ToString()))
                    return;
                oMenuPackage =(SAPbouiCOM.MenuCreationParams) objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuPackage.Image = ImagePath;
                oMenuPackage.Position = Position;
                oMenuPackage.Type = MenuType;
                oMenuPackage.UniqueID = UniqueID;
                oMenuPackage.String = DisplayName;
                parentmenu.SubMenus.AddEx(oMenuPackage);
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            // Return ParentMenu.SubMenus.Item(UniqueID)
        }

        #endregion

        public bool FormExist(string FormID)
        {
            bool FormExistRet = false;
            try
            {
                FormExistRet = false;
                foreach (SAPbouiCOM.Form uid in clsModule.objaddon.objapplication.Forms)
                {
                    if (uid.TypeEx == FormID)
                    {
                        FormExistRet = true;
                        break;
                    }
                }
                if (FormExistRet)
                {
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Visible = true;
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Select();
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return FormExistRet;

        }

        #region VIRTUAL FUNCTIONS
        public virtual void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        { }

        public virtual void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        { }

        public virtual void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo oEventInfo, ref bool BubbleEvent)
        { }        

        public virtual void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        { }

      
        #endregion

        #region ItemEvent

        private void objapplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)  
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                {
                    case clsAPInvoice.formtype:
                        aPInvoice.ItemEvent(FormUID, ref pVal,ref BubbleEvent);
                        break;
                    case clsGRPO.formtype:
                        gRPO.ItemEvent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case clsDocumentSettings.formtype:
                        documentSettings.ItemEvent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                }
               
                if (pVal.BeforeAction)
                {
                    {
                        //objform = objapplication.Forms.Item(FormUID);
                        switch (pVal.EventType)
                        {
                            case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                                {
                                    //if (pVal.FormTypeEx != "425") return;
                                    //objform = clsModule.objaddon.objapplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                                    //if (udfForm.Items.Item("U_JEDoc").Enabled == false) udfForm.Items.Item("U_JEDoc").Enabled = true;
                                    //((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String = "";
                                    //objform.Items.Item("4").Click();
                                    break;
                                }                            
                        }
                    }

                }
                else
                {
                    switch (pVal.EventType)
                    {                        
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            {
                                if (pVal.FormTypeEx == "392")
                                {
                                    objform = objapplication.Forms.Item(FormUID);
                                    SAPbouiCOM.Matrix Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("76").Specific;
                                    Matrix0.Columns.Item("U_AT_RecNum").Editable = false;      
                                }
                                break;
                            }
                        case SAPbouiCOM.BoEventTypes.et_FORM_DRAW:
                            //if (pVal.FormTypeEx == "0") return;
                            //if (pVal.FormTypeEx != "143" ) return; //143-GRPO, 141- A/P Invoice  && pVal.FormTypeEx != "141" 
                            //objform = clsModule.objaddon.objapplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                            //if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) return;
                            //udfForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                            //if (pVal.FormTypeEx != "143" || pVal.FormTypeEx != "141")
                            //{
                            //    if (udfForm.Items.Item("U_JEDoc").Enabled == false) udfForm.Items.Item("U_JEDoc").Enabled = true;
                            //   ((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String = "";
                            //    objform.Items.Item("4").Click();
                            //}
                            ////else if (pVal.FormTypeEx != "392")
                            ////{
                            ////    ((SAPbouiCOM.EditText)udfForm.Items.Item("U_APInvDoc").Specific).String = "";
                            ////    ((SAPbouiCOM.EditText)udfForm.Items.Item("U_GRPODoc").Specific).String = "";
                            ////}
                            break;

                    }
                }
                
            }
            catch (Exception ex) {
                //objapplication.StatusBar.SetText("objapplication_ItemEvent" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
           
           
        }

        #endregion

        #region FormDataEvent

        private void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch(BusinessObjectInfo.FormTypeEx)
            {
                    case clsAPInvoice.formtype:
                        aPInvoice.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);                        
                        break;
                    case clsGRPO.formtype:
                        gRPO.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                }
            }
            catch (Exception)
            {

                //throw;
            }
            

        }

        #endregion
        
        #region MenuEvent

        private void objapplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;
            switch (pVal.MenuUID)
            {
                case "1281":
                case "1282":
                case "1283":
                case "1284":
                case "1285":
                case "1286":
                case "1287":
                case "1300":
                case "1288":
                case "1289":
                case "1290":
                case "1291":
                case "1304":
                case "1292":
                case "1293":
                    objmenuevent.MenuEvent_For_StandardMenu(ref pVal, ref BubbleEvent);
                    break;
                //case "REVRECO":
                //case "REVPRJMSTR":
                //    MenuEvent_For_FormOpening(ref pVal, ref BubbleEvent);
                //    break;

            }


        }


        public void MenuEvent_For_FormOpening(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        
                        case "":
                            {
                                //FrmProjectMaster activeform = new FrmProjectMaster();
                                //activeform.Show();
                                break;
                            }
                            
                    }

                }
            }
            catch (Exception ex)
            {
                // objaddon.objapplication.SetStatusBarMessage("Error in Form Opening MenuEvent" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            }
        }

        #endregion

        #region RightClickEvent

        private void objapplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "":
                        objrightclickevent.RightClickEvent(ref eventInfo, ref BubbleEvent);
                    break;                    
                }
                

            }
            catch (Exception ex) { }
            
        }

        #endregion

        #region AppEvent

        private void objapplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {          
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    //objapplication.MessageBox("A Shut Down Event has been caught" + Environment.NewLine + "Terminating Add On...", 1, "Ok", "", "");
                    try
                    {
                        DisConnect_Addon();
                        //System.Windows.Forms.Application.Exit();
                        //if (objapplication != null)
                        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
                        //if (objcompany != null)
                        //{
                        //    if (objcompany.Connected)
                        //        objcompany.Disconnect();
                        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                        //}                        
                        //GC.Collect();                        
                        ////Environment.Exit(0);
                    }
                    catch (Exception ex)
                    {
                    }               
                    break;
               
            }
        }

        private void DisConnect_Addon()
        {
            try
            {
                if (clsModule.objaddon.objapplication.Forms.Count > 0)
                {
                    try
                    {
                        for (int frm = clsModule.objaddon.objapplication.Forms.Count - 1; frm >= 0; frm--)
                        {
                            if (clsModule.objaddon.objapplication.Forms.Item(frm).IsSystem == true) continue;
                            clsModule.objaddon.objapplication.Forms.Item(frm).Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                if (objcompany.Connected)
                    objcompany.Disconnect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
                objcompany = null;
                GC.Collect();
                System.Windows.Forms.Application.Exit();
                //Environment.Exit(0);
            }
            catch (Exception ex)
            {

            }
        }

        private void Remove_Menu(string[] MenuID = null)
        {
            try
            {
                string[] split_char;
                if (MenuID != null)
                {
                    if (MenuID.Length > 0)
                    {
                        for (int i = 0, loopTo = MenuID.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(MenuID[i]))
                                continue;
                            split_char = MenuID[i].Split(Convert.ToChar(","));
                            if (split_char.Length != 2)
                                continue;
                            if (clsModule.objaddon.objapplication.Menus.Item(split_char[0]).SubMenus.Exists(split_char[1]))
                                clsModule.objaddon.objapplication.Menus.Item(split_char[0]).SubMenus.RemoveEx(split_char[1]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }

        }

        #endregion


    }


}

