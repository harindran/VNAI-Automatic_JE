using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Automatic_JE.Common;
using SAPbobsCOM;

namespace Automatic_JE.Business_Objects
{
    class clsGRPO
    {
        public const string formtype = "143", UDFFormtype = "-143";
        private SAPbouiCOM.Form objform,udfForm;
        private string  strQuery, strSQL, Localization, TransId = "";
        //SAPbouiCOM.Matrix Matrix0;
        private bool tranflag =false;
        SAPbobsCOM.Recordset Recordset;
        SAPbouiCOM.DBDataSource odbdsContent;
        int errorCode;
        SAPbouiCOM.Button oButton;

        public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                objform =clsModule. objaddon.objapplication.Forms.Item(FormUID);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            if (pVal.ItemUID == "1" && objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                //clsModule.objaddon.objglobalmethods.WriteErrorLog("Helo");
                                if (((SAPbouiCOM.ComboBox)objform.Items.Item("3").Specific).Selected.Value == "I") return;
                                if (clsModule.objaddon.HANA)
                                {
                                    strQuery = "Select \"U_GLCode\" from OADM";
                                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);                                    
                                }
                                else
                                {
                                    strQuery = "Select U_GLCode from OADM";
                                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);                                    
                                }
                                if (strQuery == "")
                                {
                                    BubbleEvent = false;
                                    clsModule.objaddon.objapplication.MessageBox("G/L Code is not mapped in Document Settings...", 1, "OK");
                                    clsModule.objaddon.objapplication.StatusBar.SetText("G/L Code is not mapped in Document Settings...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return;
                                }
                            }

                            break;
                    }
                }
                else
                    switch (pVal.EventType)
                    {
                        case  SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                            try
                            {
                                if (objform.Items.Item("btnsje").UniqueID == "btnsje")
                                    objform.Items.Item("btnsje").Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5;
                            }
                            catch (Exception)
                            {
                            }
                           
                            break;
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            if (pVal.ItemUID == "btnsje" && (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE || objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                            {
                                //If Failure to create in Add Mode
                                if (objform.Items.Item(pVal.ItemUID).Enabled == false) return;
                                odbdsContent = objform.DataSources.DBDataSources.Item("OPDN");//Content
                                if (odbdsContent.GetValue("DocEntry", 0) == "") return;
                                udfForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);

                                if (((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String != "")
                                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select Case When U_GRPODoc=" + odbdsContent.GetValue("DocEntry", 0) + " Then 'TRUE' Else 'FALSE' End from OJDT Where TransId="+ ((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String  + " and StornoToTr is null ");
                                if (strQuery == "FALSE") { goto CreateJE; }
                                strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as [Status] from OPDN T0 Where T0.DocType='S' and T0.DocEntry=" + odbdsContent.GetValue("DocEntry", 0) + " and T0.U_JEDoc is not null ");
                                if (strQuery == "1") { objform.Items.Item(pVal.ItemUID).Enabled = false; clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry already Created. Please refresh the document...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return; }
                                if (clsModule.objaddon.objapplication.MessageBox("Do you want to create the Service JE. Continue?", 2, "Yes", "No") != 1) return;
                                CreateJE:
                                if (objform.Title.ToUpper() == "GOODS RECEIPT PO - CANCELLATION") return;
                                if (((SAPbouiCOM.ComboBox)objform.Items.Item("3").Specific).Selected.Value == "I") return;

                                //strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as [Status] from OPDN T0 Left Join [@ATPL_SJE] T1 On T0.DocEntry=T1.U_BaseEntry Where T0.DocType='S' and T0.DocEntry=" + odbdsContent.GetValue("DocEntry", 0) + " and T0.U_JEDoc is null and isnull(T1.U_Flag,'N')='N' ");
                                //strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as [Status] from OPDN T0 Where T0.DocType='S' and T0.DocEntry=" + odbdsContent.GetValue("DocEntry", 0) + " and T0.U_JEDoc is null ");
                                strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as [Status] from OPDN T0 Inner Join OJDT T1 On T0.DocEntry=T1.U_GRPODoc Where T0.DocType='S' and T0.DocEntry=" + odbdsContent.GetValue("DocEntry", 0) + " and T0.U_JEDoc is not null  ");
                                if (strQuery == "")
                                {                            
                                    clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    TransId = JournalEntry(objform.UniqueID, out TransId);
                                    clsModule.objaddon.objcompany.GetLastError(out errorCode, out strSQL);
                                    Service_JE_Logs(objform.UniqueID, odbdsContent.GetValue("DocEntry", 0), "20", TransId, (TransId != "") ? "Y" : "N", errorCode, strSQL); //logs
                                    if (TransId != "")
                                    {
                                        clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    }
                                    else
                                    {
                                        clsModule.objaddon.objapplication.MessageBox("Journal Entry Transaction Failed: Error Code: " + errorCode + " Error Desc: " + strSQL, 1, "OK");
                                        clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Transaction Failed: Error Code: " + errorCode + " Error Desc: " + strSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        TransId = ""; return;
                                    }
                                    Recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    if (clsModule.objaddon.HANA)
                                    {
                                        strQuery = "Update OJDT Set \"U_GRPODoc\"=" + odbdsContent.GetValue("DocEntry", 0) + " Where \"TransId\"=" + TransId + " ";
                                        Recordset.DoQuery(strQuery);
                                        strQuery = "Update OPDN Set \"U_JEDoc\"=" + TransId + " Where \"DocEntry\"=" + odbdsContent.GetValue("DocEntry", 0) + " ";
                                        Recordset.DoQuery(strQuery);
                                    }
                                    else
                                    {
                                        strQuery = "Update OJDT Set U_GRPODoc=" + odbdsContent.GetValue("DocEntry", 0) + " Where TransId=" + TransId + " ";
                                        Recordset.DoQuery(strQuery);
                                        strQuery = "Update OPDN Set U_JEDoc=" + TransId + " Where DocEntry=" + odbdsContent.GetValue("DocEntry", 0) + " ";
                                        Recordset.DoQuery(strQuery);
                                    }
                                }
                            }
                                break;
                    }
            }
            catch (Exception ex)
            {
            }
        }

        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID);
                //Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("39").Specific;
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        if (((SAPbouiCOM.ComboBox)objform.Items.Item("3").Specific).Selected.Value == "I") return;                        

                        if (BusinessObjectInfo.BeforeAction == true && BusinessObjectInfo.ActionSuccess == false)
                        {
                            try
                            {
                                if (tranflag == true) return;

                                if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();

                                TransId = JournalEntry(objform.UniqueID, out TransId);
                                if (TransId != "") { tranflag = true; }
                                else { tranflag = false; }

                                if (tranflag == true)
                                {
                                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                else
                                {
                                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    clsModule.objaddon.objcompany.GetLastError(out errorCode, out strQuery);
                                    clsModule.objaddon.objapplication.MessageBox("Rolled back Journal Entry: " + strQuery + "-" + errorCode, 1, "OK");
                                    clsModule.objaddon.objapplication.SetStatusBarMessage("Rolled back Journal Entry..." + strQuery, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                    BubbleEvent = false;

                                }
                            }
                            catch (Exception ex)
                            {
                                clsModule.objaddon.objapplication.MessageBox("Transaction Failed: " + ex.Message, 1, "OK");
                                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                            }
                            
                        }
                        else
                        {
                            if(BusinessObjectInfo.ActionSuccess == true)
                            {
                                odbdsContent = objform.DataSources.DBDataSources.Item("OPDN");//Content
                                if (odbdsContent.GetValue("DocEntry", 0) == "") return;

                                if (objform.Title.ToUpper() == "GOODS RECEIPT PO - CANCELLATION")
                                {
                                    if (odbdsContent.GetValue("U_JEDoc", 0) != "") Cancelling_Service_JournalEntry(objform.UniqueID, odbdsContent.GetValue("U_JEDoc", 0));
                                    return;
                                }
                                else
                                {
                                    if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();
                                    clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    TransId = JournalEntry(objform.UniqueID, out TransId);
                                    if (TransId != "") { tranflag = true; }
                                    else { tranflag = false; }
                                    try
                                    {
                                        if (tranflag == true)
                                        {
                                            if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                            clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                        }
                                        else
                                        {
                                            clsModule.objaddon.objcompany.GetLastError(out errorCode, out strSQL);
                                            clsModule.objaddon.objapplication.MessageBox("Journal Entry Transaction Failed: Error Code: " + errorCode + " Error Desc: " + strSQL, 1, "OK");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        clsModule.objaddon.objcompany.GetLastError(out errorCode, out strSQL);
                                        Service_JE_Logs(objform.UniqueID, odbdsContent.GetValue("DocEntry", 0), "20", TransId,  "N", errorCode, strSQL); //logs
                                        clsModule.objaddon.objapplication.MessageBox("Transaction Failed: " + ex.Message, 1, "OK");
                                        clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false; TransId = ""; tranflag = false;return;
                                    }                                    

                                    clsModule.objaddon.objcompany.GetLastError(out errorCode, out strSQL);
                                    odbdsContent = objform.DataSources.DBDataSources.Item("OPDN");//Content
                                    Service_JE_Logs(objform.UniqueID, odbdsContent.GetValue("DocEntry", 0), "20", TransId, (tranflag == true) ? "Y" : "N", errorCode, strSQL); //logs
                                    Recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    if (clsModule.objaddon.HANA)
                                    {
                                        strQuery = "Update OJDT Set \"U_GRPODoc\"=" + odbdsContent.GetValue("DocEntry", 0) + " Where \"TransId\"=" + TransId + " ";
                                        Recordset.DoQuery(strQuery);
                                        strQuery = "Update OPDN Set \"U_JEDoc\"=" + TransId + " Where \"DocEntry\"=" + odbdsContent.GetValue("DocEntry", 0) + " ";
                                        Recordset.DoQuery(strQuery);
                                    }
                                    else
                                    {
                                        strQuery = "Update OJDT Set U_GRPODoc=" + odbdsContent.GetValue("DocEntry", 0) + " Where TransId=" + TransId + " ";
                                        Recordset.DoQuery(strQuery);
                                        strQuery = "Update OPDN Set U_JEDoc=" + TransId + " Where DocEntry=" + odbdsContent.GetValue("DocEntry", 0) + " ";
                                        Recordset.DoQuery(strQuery);
                                    }
                                    TransId = ""; tranflag = false;
                                }                               
                            }     
                            else
                            {

                            }
                        }
                        
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        break;
                   
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        
                        if (BusinessObjectInfo.BeforeAction == true)  Create_Customize_Fields(objform.UniqueID);
                        if (objform.UDFFormUID!="")
                        {
                            udfForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                            odbdsContent = objform.DataSources.DBDataSources.Item("OPDN");//Content
                            //if (((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String != "") strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select Case When U_GRPODoc is not null then 'GRPO' Else 'A/P INVOICE' End from OJDT Where TransId=" + ((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String + " and StornoToTr is null");
                            if (((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String != "") strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select Case When U_GRPODoc=" + odbdsContent.GetValue("DocEntry", 0) + " Then 'TRUE' Else 'FALSE' End from OJDT Where TransId=" + ((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String + " and StornoToTr is null ");                            
                            if (strSQL == "TRUE" ||((SAPbouiCOM.ComboBox)objform.Items.Item("3").Specific).Selected.Value == "I") objform.Items.Item("btnsje").Enabled = false;

                        }
                        else
                        {
                            objform.Items.Item("btnsje").Enabled = false;
                        }
                        break;
                        
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.MessageBox("Exception: " + ex.Message, 1, "OK");
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
        }

        private string JournalEntry(string FormUID,out string JETransId)
        {
            JETransId = "";
            try
            {
                string  Series, Branch;
                SAPbobsCOM.JournalEntries objjournalentry;
                SAPbouiCOM.EditText oEdit;
                DateTime DocDate;                
                double LineTotal=0;
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);

                odbdsContent = objform.DataSources.DBDataSources.Item("PDN1");//Content

                objjournalentry = (SAPbobsCOM.JournalEntries)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                //clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                oEdit = (SAPbouiCOM.EditText)objform.Items.Item("10").Specific; // Posting Date
                DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);

                objjournalentry.ReferenceDate = DocDate; // Posting Date
                //oEdit = objform.Items.Item("121").Specific; // Due Date
                //DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                //objjournalentry.DueDate = DocDate;   // Due Date
                oEdit = (SAPbouiCOM.EditText)objform.Items.Item("46").Specific; // Tax Date
                DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                objjournalentry.TaxDate = DocDate;   // Document Date

                //objjournalentry.Reference = "Rev Rec JE";
                //objjournalentry.Reference2 = "Rev Rec On: " + DateTime.Now.ToString();
                //objjournalentry.UserFields.Fields.Item("U_JEDoc").Value ="";
                //objjournalentry.UserFields.Fields.Item("U_RevRecDE").Value = "";
                //objjournalentry.Memo = "Service - " + objform.Title +" - "+ ((SAPbouiCOM.EditText)objform.Items.Item("4").Specific).String;
                objjournalentry.Memo = "Service - GRPO" + " - " + ((SAPbouiCOM.EditText)objform.Items.Item("4").Specific).String;
                objjournalentry.UserFields.Fields.Item("U_AT_AutoJE").Value = "Y";
                if (clsModule.objaddon.HANA)
                {
                    Localization = clsModule.objaddon.objglobalmethods.getSingleValue("select \"LawsSet\" from CINF");
                    strQuery = "Select \"BPLId\" from OBPL where \"Disabled\"='N' and \"MainBPL\"='Y'";
                    Branch = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);
                    if (Branch == "") Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 \"Series\" FROM NNM1 WHERE \"ObjectCode\"='30' and \"Indicator\"=(Select \"Indicator\" from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between \"F_RefDate\" and \"T_RefDate\") " + " and Ifnull(\"Locked\",'')='N' ");
                    else Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 \"Series\" FROM NNM1 WHERE \"ObjectCode\"='30' and \"Indicator\"=(Select \"Indicator\" from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between \"F_RefDate\" and \"T_RefDate\") " + " and Ifnull(\"Locked\",'')='N' and \"BPLId\"='" + Branch + "'");
                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"U_GLCode\" from OADM");

                }
                else
                {
                    Localization = clsModule.objaddon.objglobalmethods.getSingleValue("select LawsSet from CINF");
                    strQuery = "Select BPLId from OBPL where Disabled='N' and MainBPL='Y'";
                    Branch = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);
                    if (Branch == "") Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between F_RefDate and T_RefDate) " + " and Isnull(Locked,'')='N' ");
                    else Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between F_RefDate and T_RefDate) " + " and Isnull(Locked,'')='N' and BPLId='" + Branch + "'");
                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select U_GLCode from OADM");

                }
                if (Localization != "IN")
                {
                    objjournalentry.AutoVAT = BoYesNoEnum.tNO;
                    objjournalentry.AutomaticWT = BoYesNoEnum.tNO;
                }
                if (!string.IsNullOrEmpty(Series)) objjournalentry.Series = Convert.ToInt32(Series);
                

                for (int ContentRow = 0; ContentRow <= odbdsContent.Size - 1; ContentRow++)
                {
                    if (odbdsContent.GetValue("AcctCode", ContentRow) != "") //odbdsContent.GetValue("Dscription", ContentRow)
                    {
                        objjournalentry.Lines.AccountCode = Convert.ToString(odbdsContent.GetValue("AcctCode", ContentRow));
                        objjournalentry.Lines.Debit = Convert.ToDouble(odbdsContent.GetValue("LineTotal", ContentRow));
                        if (Branch != "") objjournalentry.Lines.BPLID = Convert.ToInt32(Branch);
                        if (Convert.ToString(odbdsContent.GetValue("OcrCode", ContentRow)) != "") objjournalentry.Lines.CostingCode = Convert.ToString(odbdsContent.GetValue("OcrCode", ContentRow));
                        if (Convert.ToString(odbdsContent.GetValue("OcrCode2", ContentRow)) != "") objjournalentry.Lines.CostingCode2 = Convert.ToString(odbdsContent.GetValue("OcrCode2", ContentRow));
                        if (Convert.ToString(odbdsContent.GetValue("OcrCode3", ContentRow)) != "") objjournalentry.Lines.CostingCode3 = Convert.ToString(odbdsContent.GetValue("OcrCode3", ContentRow));
                        if (Convert.ToString(odbdsContent.GetValue("OcrCode4", ContentRow)) != "") objjournalentry.Lines.CostingCode4 = Convert.ToString(odbdsContent.GetValue("OcrCode4", ContentRow));
                        if (Convert.ToString(odbdsContent.GetValue("OcrCode5", ContentRow)) != "") objjournalentry.Lines.CostingCode5 = Convert.ToString(odbdsContent.GetValue("OcrCode5", ContentRow));
                        objjournalentry.Lines.Add();
                        //objjournalentry.Lines.SetCurrentLine(ContentRow);
                        LineTotal+= Convert.ToDouble(odbdsContent.GetValue("LineTotal", ContentRow));
                    }
                }
                
                objjournalentry.Lines.AccountCode = strQuery;
                objjournalentry.Lines.Credit = Convert.ToDouble(LineTotal);
                if (Branch != "") objjournalentry.Lines.BPLID = Convert.ToInt32(Branch);
                //if (Convert.ToString(odbdsContent.GetValue("OcrCode", 0)) != "") objjournalentry.Lines.CostingCode = Convert.ToString(odbdsContent.GetValue("OcrCode", 0));
                //if (Convert.ToString(odbdsContent.GetValue("OcrCode2", 0)) != "") objjournalentry.Lines.CostingCode2 = Convert.ToString(odbdsContent.GetValue("OcrCode2", 0));
                //if (Convert.ToString(odbdsContent.GetValue("OcrCode3", 0)) != "") objjournalentry.Lines.CostingCode3 = Convert.ToString(odbdsContent.GetValue("OcrCode3", 0));
                //if (Convert.ToString(odbdsContent.GetValue("OcrCode4", 0)) != "") objjournalentry.Lines.CostingCode4 = Convert.ToString(odbdsContent.GetValue("OcrCode4", 0));
                //if (Convert.ToString(odbdsContent.GetValue("OcrCode5", 0)) != "") objjournalentry.Lines.CostingCode5 = Convert.ToString(odbdsContent.GetValue("OcrCode5", 0));
                objjournalentry.Lines.Add();

                
                if (objjournalentry.Add() != 0)
                {
                    clsModule.objaddon.objapplication.MessageBox("Journal Transaction: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(),1, "OK");
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Journal: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry);
                    return JETransId;
                }
                else
                {                    
                    strQuery = clsModule.objaddon.objcompany.GetNewObjectKey();
                    JETransId = strQuery;
                    //udfForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    //((SAPbouiCOM.EditText)udfForm.Items.Item("U_JEDoc").Specific).String = TransId;
                    //clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    return JETransId;
                }
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.MessageBox("JE Posting Error: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), 1, "OK");
                clsModule.objaddon.objapplication.SetStatusBarMessage("JE Posting Error: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return JETransId;

            }

        }

        private bool Cancelling_Service_JournalEntry(string FormUID, string JETransId)
        {
            string TransId;
            SAPbobsCOM.JournalEntries objjournalentry;
            if (string.IsNullOrEmpty(JETransId)) return true;
            SAPbobsCOM.Recordset objRs;
            string strSQL;
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                odbdsContent = objform.DataSources.DBDataSources.Item("OPDN");//Content
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (clsModule.objaddon.HANA)
                {
                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("select distinct 1 as \"Status\" from OJDT where \"StornoToTr\"=" + JETransId + "");
                }
                else
                {
                    strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("select distinct 1 as Status from OJDT where StornoToTr=" + JETransId + "");
                }
                    
                if (strQuery == "1")
                {
                    return true;
                    //TransId = clsModule.objaddon.objglobalmethods.getSingleValue("select \"TransId\" from OJDT where \"StornoToTr\"=" + JETransId + "");
                }
                if (clsModule.objaddon.HANA)
                {
                    strSQL = "Select T0.\"Series\",T0.\"TaxDate\",T0.\"DueDate\",T0.\"RefDate\",T0.\"Ref1\",T0.\"Ref2\",T0.\"Memo\",T1.\"Account\",T1.\"Credit\",T1.\"Debit\",T1.\"BPLId\",";
                    strSQL += "\n (Select \"CardCode\" from OCRD where \"CardCode\"=T1.\"ShortName\") as \"BPCode\",\"ProfitCode\",\"OcrCode2\",\"OcrCode3\",\"OcrCode4\",\"OcrCode5\"";
                    strSQL += "\n from OJDT T0 join JDT1 T1 ON T0.\"TransId\"=T1.\"TransId\" where T1.\"TransId\"='" + JETransId + "' order by T1.\"Line_ID\"";
                }
                 else
                {
                    strSQL = "Select T0.Series,T0.TaxDate,T0.DueDate,T0.RefDate,T0.Ref1,T0.Ref2,T0.Memo,T1.Account,T1.Credit,T1.Debit,T1.BPLId,";
                    strSQL += "\n (Select CardCode from OCRD where CardCode=T1.ShortName) as BPCode,ProfitCode,OcrCode2,OcrCode3,OcrCode4,OcrCode5";
                    strSQL += "\n from OJDT T0 join JDT1 T1 ON T0.TransId=T1.TransId where T1.TransId='" + JETransId + "' order by T1.Line_ID";
                }
                objRs.DoQuery(strSQL);
                if (objRs.RecordCount == 0) return true;
                if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();
                objjournalentry = (SAPbobsCOM.JournalEntries)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                clsModule.objaddon.objapplication.StatusBar.SetText("Service Journal Entry Reversing Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                objjournalentry.TaxDate = Convert.ToDateTime(objRs.Fields.Item("TaxDate").Value); 
                objjournalentry.DueDate = Convert.ToDateTime(objRs.Fields.Item("DueDate").Value); 
                objjournalentry.ReferenceDate = Convert.ToDateTime(objRs.Fields.Item("RefDate").Value); 
                objjournalentry.Reference = Convert.ToString(objRs.Fields.Item("Ref1").Value); 
                objjournalentry.Reference2 = Convert.ToString(objRs.Fields.Item("Ref2").Value);
                          
                
                //objjournalentry.Reference3 = DateTime.Now.ToString();
                //objjournalentry.Memo = Convert.ToString(objRs.Fields.Item("Memo").Value) + "(Reversal) - " + JETransId;
                objjournalentry.Memo = "Service - GRPO" + " - " + ((SAPbouiCOM.EditText)objform.Items.Item("4").Specific).String + " (Reversal) - " + JETransId;
                //objjournalentry.Series = Convert.ToInt32(objRs.Fields.Item("Series").Value); 
                objjournalentry.UserFields.Fields.Item("U_AT_AutoJE").Value = "N";
                for (int AccRow = 0; AccRow < objRs.RecordCount; AccRow++)
                {
                    if (Convert.ToString(objRs.Fields.Item("BPCode").Value) != "")
                        objjournalentry.Lines.ShortName = Convert.ToString(objRs.Fields.Item("BPCode").Value);
                    else
                        objjournalentry.Lines.AccountCode = Convert.ToString(objRs.Fields.Item("Account").Value);
                    if (Convert.ToDouble(objRs.Fields.Item("Credit").Value) != 0)
                        objjournalentry.Lines.Debit = Convert.ToDouble(objRs.Fields.Item("Credit").Value);
                    else
                        objjournalentry.Lines.Credit = Convert.ToDouble(objRs.Fields.Item("Debit").Value);
                    if (Convert.ToString(objRs.Fields.Item("BPLId").Value) != "") objjournalentry.Lines.BPLID = Convert.ToInt32(objRs.Fields.Item("BPLId").Value);
                    if (Convert.ToString(objRs.Fields.Item("ProfitCode").Value) != "") objjournalentry.Lines.CostingCode = Convert.ToString(objRs.Fields.Item("ProfitCode").Value);
                    if (Convert.ToString(objRs.Fields.Item("OcrCode2").Value) != "") objjournalentry.Lines.CostingCode2 = Convert.ToString(objRs.Fields.Item("OcrCode2").Value);
                    if (Convert.ToString(objRs.Fields.Item("OcrCode3").Value) != "") objjournalentry.Lines.CostingCode3 = Convert.ToString(objRs.Fields.Item("OcrCode3").Value);
                    if (Convert.ToString(objRs.Fields.Item("OcrCode4").Value) != "") objjournalentry.Lines.CostingCode4 = Convert.ToString(objRs.Fields.Item("OcrCode4").Value);
                    if (Convert.ToString(objRs.Fields.Item("OcrCode5").Value) != "") objjournalentry.Lines.CostingCode5 = Convert.ToString(objRs.Fields.Item("OcrCode5").Value);

                    
                    objjournalentry.Lines.Add();
                    objRs.MoveNext();
                }

                if (objjournalentry.Add() != 0)
                {
                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    clsModule.objaddon.objapplication.MessageBox("Auto-Cancellation failed. Please cancel the Service Journal Entry Manually!" , 1, "OK");
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Reverse: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry);
                    return false;
                }
                
                else
                {
                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    TransId = clsModule.objaddon.objcompany.GetNewObjectKey();
                    if (clsModule.objaddon.HANA)
                    {
                        objRs.DoQuery("Update OJDT set \"StornoToTr\"=" + JETransId + " where \"TransId\"=" + TransId + "");
                        objRs.DoQuery("Update OJDT set \"U_RevJEDoc\"=" + TransId + ",\"U_AT_AutoJE\"='N' where \"TransId\"=" + JETransId + "");
                    }
                    else
                    {
                        objRs.DoQuery("Update OJDT set StornoToTr=" + JETransId + " where TransId=" + TransId + "");
                        objRs.DoQuery("Update OJDT set U_RevJEDoc=" + TransId + ",U_AT_AutoJE='N' where TransId=" + JETransId + "");
                    }
                    //******Reconciling the Cancelled & Reversed JE************
                    if(Account_Reconciliation(JETransId, TransId) == false) { }
                    clsModule.objaddon.objapplication.StatusBar.SetText("Service Journal Entry Reversed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                }

               
            }
            catch (Exception ex)
            {
                if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                clsModule.objaddon.objapplication.SetStatusBarMessage("Transaction Cancelling Error " + clsModule.objaddon.objcompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;
            }


        }

        private bool Account_Reconciliation(string CanGRPOJE, string RevGRPOJE)
        {
            try
            {
                SAPbobsCOM.Recordset objRs, updateRset;
                Recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                updateRset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                double RecAmount = 0;
                string JE = string.Concat(CanGRPOJE, ",", RevGRPOJE);

                if (clsModule.objaddon.HANA)
                {
                    strQuery = "Select T0.\"TransId\",T0.\"Line_ID\",\"T0.Account\",Case when T0.\"DebCred\" = 'D' Then T0.\"BalDueDeb\" Else T0.\"BalDueCred\" End \"Amount\" from JDT1 T0 Join OJDT T1 On T0.\"TransId\" = T1.\"TransId\" Where T0.\"TransId\" in (" + CanGRPOJE + ") and (T0.\"BalDueDeb\"<>0 or T0.\"BalDueCred\"<>0)";
                }
                else
                {
                    strQuery = "Select T0.TransId,T0.Line_ID,T0.Account,Case when T0.DebCred = 'D' Then T0.BalDueDeb Else T0.BalDueCred End Amount from JDT1 T0 Join OJDT T1 On T0.TransId = T1.TransId Where T0.TransId in (" + CanGRPOJE + ") and (T0.BalDueDeb<>0 or T0.BalDueCred<>0)";
                }
                Recordset.DoQuery(strQuery);
                if (Recordset.RecordCount > 0)
                {
                    for (int DTRow = 0; DTRow <= Recordset.RecordCount - 1; DTRow++)
                    {
                        if (clsModule.objaddon.HANA)
                        {
                            strQuery = "Select T0.\"TransId\",T0.\"Line_ID\",\"T0.Account\",Case when T0.\"DebCred\" = 'D' Then T0.\"BalDueDeb\" Else T0.\"BalDueCred\" End \"Amount\" from JDT1 T0 Join OJDT T1 On T0.\"TransId\" = T1.\"TransId\" Where T0.\"TransId\" in (" + JE + ") and T0.\"Account\"='" + Convert.ToString(Recordset.Fields.Item("Account").Value) + "' and T0.\"Line_ID\"='" + Convert.ToString(Recordset.Fields.Item("Line_ID").Value) + "'";
                            strSQL = "Select MIN(Case when T0.\"DebCred\" = 'D' Then T0.\"BalDueDeb\" Else T0.\"BalDueCred\" End) \"RecAmount\" from JDT1 T0 Join OJDT T1 On T0.\"TransId\" = T1.\"TransId\" Where T0.\"TransId\" in (" + JE + ") and T0.\"Account\" = '" + Convert.ToString(Recordset.Fields.Item("Account").Value) + "' and T0.\"Line_ID\"='" + Convert.ToString(Recordset.Fields.Item("Line_ID").Value) + "'";
                        }
                        else
                        {
                            strQuery = "Select T0.TransId,T0.Line_ID,T0.Account,Case when T0.DebCred = 'D' Then T0.BalDueDeb Else T0.BalDueCred End Amount from JDT1 T0 Join OJDT T1 On T0.TransId = T1.TransId Where T0.TransId in (" + JE + ") and T0.Account='" + Convert.ToString(Recordset.Fields.Item("Account").Value) + "' and T0.Line_ID='" + Convert.ToString(Recordset.Fields.Item("Line_ID").Value) + "'";
                            strSQL = "Select MIN(Case when T0.DebCred = 'D' Then T0.BalDueDeb Else T0.BalDueCred End) RecAmount from JDT1 T0 Join OJDT T1 On T0.TransId = T1.TransId Where T0.TransId in (" + JE + ") and T0.Account = '" + Convert.ToString(Recordset.Fields.Item("Account").Value) + "' and T0.Line_ID='" + Convert.ToString(Recordset.Fields.Item("Line_ID").Value) + "'";
                        }
                        objRs.DoQuery(strQuery);
                        if (objRs.RecordCount == 0) continue;
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue(strSQL);
                        RecAmount = Convert.ToDouble(strSQL);
                        DateTime DocDate = DateTime.ParseExact(DateTime.Now.ToString("yyyyMMdd"), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                        IInternalReconciliationsService service = (IInternalReconciliationsService)clsModule.objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService);
                        InternalReconciliationOpenTrans openTrans = (InternalReconciliationOpenTrans)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans);
                        IInternalReconciliationParams reconParams = (IInternalReconciliationParams)service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams);
                        int Row = 0;
                        openTrans.CardOrAccount = CardOrAccountEnum.coaAccount;
                        openTrans.ReconDate = DocDate;
                        for (int RecRow = 0; RecRow <= objRs.RecordCount - 1; RecRow++)
                        {
                            openTrans.InternalReconciliationOpenTransRows.Add();
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES;
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = Convert.ToInt32(objRs.Fields.Item("TransId").Value);
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = Convert.ToInt32(objRs.Fields.Item("Line_ID").Value);
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount;// Convert.ToDouble(objRs.Fields.Item("Amount").Value);
                            Row += 1;
                            objRs.MoveNext();
                        }
                        int Reconum = 0;
                        try
                        {
                            reconParams = service.Add(openTrans);
                        }
                        catch (Exception ex)
                        {
                            clsModule.objaddon.objapplication.StatusBar.SetText("Reconciled Error..." + "-" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return false;
                        }
                        Reconum = reconParams.ReconNum;
                        if (clsModule.objaddon.HANA)
                        {
                            strQuery = "Update JDT1 Set \"U_AT_RecNum\"=" + Reconum + "  Where \"TransId\" in (" + JE + ") and \"Account\" ='" + Convert.ToString(Recordset.Fields.Item("Account").Value) + "'";
                        }
                        else
                        {
                            strQuery = "Update JDT1 Set U_AT_RecNum=" + Reconum + " Where TransId in (" + JE + ") and Account='" + Convert.ToString(Recordset.Fields.Item("Account").Value) + "'";
                        }
                        updateRset.DoQuery(strQuery);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans);

                        Recordset.MoveNext();
                    }

                }
                clsModule.objaddon.objapplication.StatusBar.SetText("Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                GC.Collect();
                return true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Recon: " + "-" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        public bool Service_JE_Logs(string FormUID, string TranDocEntry,  string ObjType,string JETranId, string JEFlag, int ErrID, string ErrDesc, string RecStatus="",string ReconNum="")
        {
            try
            {
                bool Flag = false;
                string DocEntry,HTable="";
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralData;
                SAPbobsCOM.GeneralDataParams oGeneralParams;
                Recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("ATSJE");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                try
                {
                    DocEntry = clsModule.objaddon.objglobalmethods.getSingleValue("Select DocEntry from [@ATPL_SJE] where U_BaseEntry=" + TranDocEntry + " Order by DocEntry Desc");
                    oGeneralParams.SetProperty("DocEntry", DocEntry); 
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    Flag = true;
                }
                catch (Exception ex)
                {
                    Flag = false;
                }

                if (ObjType == "20") HTable = "OPDN";
                else if (ObjType == "18") HTable = "OPCH";

                if(Flag == false)
                {
                    oGeneralData.SetProperty("U_GenDate", DateTime.Now.Date);
                    oGeneralData.SetProperty("U_BaseNo", objform.DataSources.DBDataSources.Item(HTable).GetValue("DocNum", 0));
                    oGeneralData.SetProperty("U_BaseEntry", objform.DataSources.DBDataSources.Item(HTable).GetValue("DocEntry", 0));
                    oGeneralData.SetProperty("U_DocObjType", ObjType);
                }
                if (JETranId != "") oGeneralData.SetProperty("U_JETransId", JETranId);
                //if (ErrDesc!="")
                oGeneralData.SetProperty("U_ErrDesc", ErrDesc);
                oGeneralData.SetProperty("U_ErrId", Convert.ToString(ErrID));
                oGeneralData.SetProperty("U_Flag", JEFlag);
                oGeneralData.SetProperty("U_Status", (JEFlag=="Y")?"Success": (JEFlag == "C")? "Canceled": "Failure");
                oGeneralData.SetProperty("U_RecStatus", RecStatus);
                oGeneralData.SetProperty("U_ReconNum", ReconNum);

                if (Flag == true)
                {
                    oGeneralService.Update(oGeneralData);
                    //clsModule.objaddon.objapplication.StatusBar.SetText("Modified Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    return true;
                }
                else
                {
                    oGeneralParams = oGeneralService.Add(oGeneralData);
                    //clsModule.objaddon.objapplication.StatusBar.SetText("Copied Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    return true;
                }

            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Service JE Logs: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private void Create_Customize_Fields(string oFormUID)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.Item(oFormUID);
                SAPbouiCOM.Item oItem;

                oItem = objform.Items.Add("btnsje", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Service JE";
                oItem.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5;
                oItem.Top = objform.Items.Item("2").Top;
                oItem.Height = objform.Items.Item("2").Height;
                oItem.LinkTo = "10000330";
                Size Fieldsize = System.Windows.Forms.TextRenderer.MeasureText("Service JE", new Font("Arial", 12.0f));
                oItem.Width = Fieldsize.Width;
                objform.Items.Item("btnsje").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                objform.Items.Item("btnsje").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                
            }
            catch (Exception ex)
            {
            }

        }
    }
}

