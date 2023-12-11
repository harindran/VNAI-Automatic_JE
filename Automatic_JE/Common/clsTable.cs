using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_JE.Common
{
    class clsTable
    {        
        public void FieldCreation()
        {            
            AddFields("OADM", "GLCode", "G/L Code", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulChartOfAccounts);
            AddFields("OJDT", "GRPODoc", "GRPO DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulPurchaseDeliveryNotes);
            AddFields("OJDT", "APInvDoc", "A/P Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulPurchaseInvoices);
            AddFields("OPDN", "JEDoc", "JE TransID", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulJournalEntries);
            AddFields("JDT1", "AT_RecNum", "Reconciliation No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("OJDT", "RevJEDoc", "Reverse JE ID", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulJournalEntries);

            AddFields("OJDT", "AT_AutoJE", "Auto-Post JE", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "N", true);
                       
            //*********Log Table
            AddTables("ATPL_SJE", "Serv JE Logs", SAPbobsCOM.BoUTBTableType.bott_Document);
            
            AddFields("@ATPL_SJE", "JETransId", "JE Trans ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@ATPL_SJE", "GenDate", "Generation Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@ATPL_SJE", "BaseNo", "Base DocNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@ATPL_SJE", "BaseEntry", "Base DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);          
            AddFields("@ATPL_SJE", "DocObjType", "Doc Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@ATPL_SJE", "ErrDesc", "Error Description", SAPbobsCOM.BoFieldTypes.db_Memo);
            AddFields("@ATPL_SJE", "ErrId", "Error ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@ATPL_SJE", "Flag", "Flag", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@ATPL_SJE", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@ATPL_SJE", "RecStatus", "Reconciliation Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@ATPL_SJE", "ReconNum", "Reconciliation Num", SAPbobsCOM.BoFieldTypes.db_Memo);

            AddUDO("ATSJE", "Service JE", SAPbobsCOM.BoUDOObjType.boud_Document, "ATPL_SJE", new[] { "" }, new[] { "DocEntry", "DocNum" }, true, true);
            //************
        }
        #region Master Data Creation

        public void Project_Master()
        {
            AddTables("AT_PROJMSTR", "Project Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddTables("AT_PROJMSTR1", "Project Master Contents", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PROJMSTR2", "Project Master Attachments", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PROJMSTR3", "BOQ Item Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PROJMSTR4", "BOQ Labour Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            
            //Header Table
            AddFields("@AT_PROJMSTR", "PrjCode", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR", "PrjName", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);       

          

            AddUDO("AT_PROJMASTER", "Revenue Project Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PROJMSTR", new[] { "AT_PROJMSTR1", "AT_PROJMSTR2", "AT_PROJMSTR3", "AT_PROJMSTR4" }, new[] { "Code", "Name" },false, true, false);
        }

        #endregion


        #region Table Creation Common Functions

        private void AddTables(string strTab, string strDesc, SAPbobsCOM.BoUTBTableType nType)
        {
            // var oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                // Adding Table
                if (!oUserTablesMD.GetByKey(strTab))
                {
                    oUserTablesMD.TableName = strTab;
                    oUserTablesMD.TableDescription = strDesc;
                    oUserTablesMD.TableType = nType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        throw new Exception(clsModule.objaddon.objcompany.GetLastErrorDescription() + strTab);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddFields(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10, SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedSysObject= 0,string UDTTable="", SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "", bool Yesno = false, string[] Validvalues = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            
            try
            {
                // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                // If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                // strTab = "@" + strTab
                // End If
                if (!IsColumnExists(strTab, strCol))
                {
                    // If Not oUserFieldMD1 Is Nothing Then
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    // End If
                    // oUserFieldMD1 = Nothing
                    // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;
                    
                    if (Yesno == true)
                    {
                        oUserFieldMD1.ValidValues.Value = "Y";
                        oUserFieldMD1.ValidValues.Description = "Yes";
                        oUserFieldMD1.ValidValues.Add();
                        oUserFieldMD1.ValidValues.Value = "N";
                        oUserFieldMD1.ValidValues.Description = "No";
                        oUserFieldMD1.ValidValues.Add();
                    }
                    if (LinkedSysObject != 0)
                        oUserFieldMD1.LinkedSystemObject = LinkedSysObject;// SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulInvoices ;
                    if (UDTTable != "")
                        oUserFieldMD1.LinkedTable = UDTTable;
                    string[] split_char;
                    if (Validvalues !=null)
            {
                        if (Validvalues.Length > 0)
                        {
                            for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                            {
                                if (string.IsNullOrEmpty(Validvalues[i]))
                                    continue;
                                split_char = Validvalues[i].Split(Convert.ToChar(","));
                                if (split_char.Length != 2)
                                    continue;
                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }
                    int val;
                    val = oUserFieldMD1.Add();
                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule. objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short,true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private bool IsColumnExists(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet=null;
            string strSQL;
            try
            {
                if (clsModule. objaddon.HANA)
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";
                }
                else
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" + Table + "' AND AliasID = '" + Column + "'";
                }

                oRecordSet = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(strSQL);

                if (Convert.ToInt32( oRecordSet.Fields.Item(0).Value) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddKey(string strTab, string strColumn, string strKey, int i)
        {
            var oUserKeysMD = default(SAPbobsCOM.UserKeysMD);

            try
            {
                // // The meta-data object must be initialized with a
                // // regular UserKeys object
                oUserKeysMD =(SAPbobsCOM.UserKeysMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

                if (!oUserKeysMD.GetByKey("@" + strTab, i))
                {

                    // // Set the table name and the key name
                    oUserKeysMD.TableName = strTab;
                    oUserKeysMD.KeyName = strKey;

                    // // Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn;
                    oUserKeysMD.Elements.Add();
                    oUserKeysMD.Elements.ColumnAlias = "RentFac";

                    // // Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

                    // // Add the key
                    if (oUserKeysMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                oUserKeysMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void AddUDO(string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType, string strTable, string[] childTable, string[] sFind, bool Cancel = false, bool canlog = false, bool Manageseries = false)
        {

           SAPbobsCOM.UserObjectsMD oUserObjectMD=null;
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
               
                if (!oUserObjectMD.GetByKey(strUDO)) //(oUserObjectMD.GetByKey(strUDO) == 0)
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;

                    if(Cancel)
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (Manageseries)
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (canlog)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + strTable.ToString();
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUserObjectMD.LogTableName = "";
                    }

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.ExtensionName = "";

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    tablecount = 1;
                    if (sFind.Length > 0)
                    {
                        for (int i = 0, loopTo = sFind.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(sFind[i]))
                                continue;
                            oUserObjectMD.FindColumns.ColumnAlias = sFind[i];
                            oUserObjectMD.FindColumns.Add();
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount);
                            tablecount = tablecount + 1;
                        }
                    }

                    tablecount = 0;
                    if (childTable != null)
            {
                        if (childTable.Length > 0)
                        {
                            for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(childTable[i]))
                                    continue;
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                oUserObjectMD.ChildTables.TableName = childTable[i];
                                oUserObjectMD.ChildTables.Add();
                                tablecount = tablecount + 1;
                            }
                        }
                    }

                    if (oUserObjectMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }


        #endregion

    }
}
