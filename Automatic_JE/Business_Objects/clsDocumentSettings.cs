using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Automatic_JE.Common;

namespace Automatic_JE.Business_Objects
{
    class clsDocumentSettings
    {
        public const string formtype = "228", UDFFormtype = "-228";
        private SAPbouiCOM.Form objform;

        public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                            {
                                break;
                            }
                    }
                }
                else
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                           
                                //Create_Controls();
                                break;
                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            //if (pVal.ItemUID == "tglcode")
                            //{
                            //    SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            //    string sCFL_ID = oCFLEvento.ChooseFromListUID;
                            //    SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item(sCFL_ID);
                            //    if (oCFLEvento.BeforeAction == false)
                            //    {
                            //        SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                            //        try
                            //        {
                            //            ((SAPbouiCOM.EditText)objform.Items.Item("tglcode").Specific).String = Convert.ToString(oDataTable.GetValue("AcctCode", 0));
                            //        }
                            //        catch (Exception ex)
                            //        {
                            //        }

                            //    }
                            //} 
                            break;
                            
                    }
            }
            catch (Exception ex)
            {
            }
        }

        private void Create_Controls()
        {
            try
            {
                SAPbouiCOM.EditText oedit;
                SAPbouiCOM.StaticText staticText;
                SAPbouiCOM.Item oitem;

                oitem = objform.Items.Add("lglcode", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oitem.Left = objform.Items.Item("6").Left + objform.Items.Item("6").Width + 10;
                oitem.Width = 60;
                oitem.Top = objform.Items.Item("6").Top + objform.Items.Item("6").Height + 5;
                oitem.Height = objform.Items.Item("6").Height;
                staticText = (SAPbouiCOM.StaticText)oitem.Specific;
                staticText.Caption = "G/L Code";
                staticText.Item.FromPane = 2;
                staticText.Item.ToPane = 2;

                oitem = objform.Items.Add("tglcode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oitem.Left = objform.Items.Item("lglcode").Left + objform.Items.Item("lglcode").Width + 5;
                oitem.Width = 100;
                oitem.Top = objform.Items.Item("lglcode").Top;
                oitem.Height = objform.Items.Item("lglcode").Height;
                oitem.LinkTo = "lglcode";
                oedit = (SAPbouiCOM.EditText)oitem.Specific;
                //oedit.Item.Enabled = false;
                oedit.DataBind.SetBound(true, "OADM", "U_GLCode");
                oedit.Item.FromPane = 2;
                oedit.Item.ToPane = 2;
                AddChooseFromList();
                oedit.ChooseFromListUID = "cflcode";
                oedit.ChooseFromListAlias = "AcctCode";


            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void AddChooseFromList()
        {
            try
            {

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                oCFLs = objform.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(clsModule.objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "1";
                oCFLCreationParams.UniqueID = "cflcode";

                oCFL = oCFLs.Add(oCFLCreationParams);

                //  Adding Conditions to CFL1

                oCons = oCFL.GetConditions();

                oCon = oCons.Add();
                oCon.Alias = "Postable";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.Alias = "LocManTran";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "N";

                oCFL.SetConditions(oCons);

               
            }
            catch
            {
            }
        }

    }
}
