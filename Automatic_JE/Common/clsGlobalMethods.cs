using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Automatic_JE.Common
{
    class clsGlobalMethods
    {
        string strsql, BankFileName;
        SAPbobsCOM.Recordset objrs;

        public string GetDocNum(string sUDOName, int Series)
        {
            string GetDocNumRet = "";
            string StrSQL;
            SAPbobsCOM.Recordset objRS;
            objRS = (SAPbobsCOM.Recordset) clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            // If objAddOn.HANA Then
            if (Series == 0)
            {
                StrSQL = " select  \"NextNumber\"  from NNM1 where \"ObjectCode\"='" + sUDOName + "'";
            }
            else
            {
                StrSQL = " select  \"NextNumber\"  from NNM1 where \"ObjectCode\"='" + sUDOName + "' and \"Series\" = " + Series;
            }
            // Else
            // StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
            // End If
            objRS.DoQuery(StrSQL);
            objRS.MoveFirst();
            if (!objRS.EoF)
            {
                return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString()).ToString();
            }
            else
            {
                GetDocNumRet = "1";
            }

            return GetDocNumRet;
        }

        public string GetNextCode_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                if (clsModule.objaddon.HANA)
                {
                    strsql = "select IFNULL(Max(CAST(\"Code\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                }
                else
                {
                    strsql = "select ISNULL(Max(CAST(Code As integer)),0)+1 from " + Tablename.ToString() + "";
                }

                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value) ;
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public string GetNextDocNum_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                strsql = "select IFNULL(Max(CAST(\"DocNum\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                objrs =(SAPbobsCOM.Recordset) clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
               clsModule. objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public string GetNextDocEntry_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                strsql = "select IFNULL(Max(CAST(\"DocEntry\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                objrs =(SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public string Convert_String_TimeHHMM(string str)
        {
            str = "0000" + Regex.Replace(str, @"[^\d]", "");
            return str.PadRight(4);
        }

        public string GetDuration_BetWeenTime(string strFrom, string strTo)
        {
            DateTime Fromtime, Totime;
            TimeSpan Duration;
            strFrom = Convert_String_TimeHHMM(strFrom);
            strTo = Convert_String_TimeHHMM(strTo);
            Totime = new DateTime(2000, 1, 1,Convert.ToInt32(strTo.PadLeft(2)), Convert.ToInt32(strTo.PadLeft(2)), 0);
            Fromtime = new DateTime(2000, 1, 1, Convert.ToInt32(strFrom.PadLeft(2)), Convert.ToInt32(strFrom.PadRight(2)), 0);
            if (Totime < Fromtime)
                Totime = new DateTime(2000, 1, 2, Convert.ToInt32(strTo.PadLeft(2)), Convert.ToInt32(strTo.PadLeft(2)), 0);
            Duration = Totime - Fromtime;
            return Duration.Hours.ToString() + "." + Duration.Minutes.ToString() + "00".PadLeft(2);
        }

        public string getSingleValue(string StrSQL)
        {
            try
            {
                SAPbobsCOM.Recordset rset =(SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strReturnVal = "";
                rset.DoQuery(StrSQL);
                return Convert.ToString((rset.RecordCount) > 0 ? rset.Fields.Item(0).Value.ToString() : "");
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(" Get Single Value Function Failed :  " + ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return "";
            }
        }

        public void LoadSeries(SAPbouiCOM.Form objform, SAPbouiCOM.DBDataSource DBSource, string ObjectType)
        {
            try
            {
                SAPbouiCOM.ComboBox ComboBox0;
                ComboBox0 =(SAPbouiCOM.ComboBox) objform.Items.Item("series").Specific;
                ComboBox0.ValidValues.LoadSeries(ObjectType, SAPbouiCOM.BoSeriesMode.sf_Add);
                ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                DBSource.SetValue("DocNum", 0, clsModule.objaddon.objglobalmethods.GetDocNum(ObjectType, Convert.ToInt32(ComboBox0.Selected.Value)));
            }
            catch (Exception ex)
            {

            }
        }

        public void WriteErrorLog(string Str)
        {
            string Foldername, Attachpath;
            Attachpath = @"C:\ProgramData\Altrocks Tech\" + clsModule.objaddon.addonName + @"\Add-on Logs\"; // getSingleValue("select ""AttachPath"" from OADP")
            Foldername = Attachpath ;
            
            if (!Directory.Exists(Foldername))            
            {
                Directory.CreateDirectory(Foldername);
            }

            FileStream fs;
            string chatlog = Foldername +  DateTime.Now.ToString("yyyy-MM-dd")+ "_" + clsModule.objaddon.objcompany.UserName + ".txt";
            if (File.Exists(chatlog))
            {
            }
            else
            {
                fs = new FileStream(chatlog, FileMode.Create, FileAccess.Write);
                fs.Close();
            }
            string sdate;
            sdate = Convert.ToString(DateTime.Now);
            if (File.Exists(chatlog) == true)
            {
                var objWriter = new StreamWriter(chatlog, true);
                objWriter.WriteLine(sdate + " : " + Str);
                objWriter.Close();
            }
            else
            {
                var objWriter = new StreamWriter(chatlog, false);
            }
        }

        public void RemoveLastrow(SAPbouiCOM.Matrix omatrix, string Columname_check)
        {
            try
            {
                if (omatrix.VisualRowCount == 0)
                    return;
                if (string.IsNullOrEmpty(Columname_check.ToString()))
                    return;
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific).String=="")
                {
                    omatrix.DeleteRow(omatrix.VisualRowCount);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void SetAutomanagedattribute_Editable(SAPbouiCOM.Form oform, string fieldname, bool add, bool find, bool update)
        {

            if (add == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (find == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (update)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }
        }

        public void SetAutomanagedattribute_Visible(SAPbouiCOM.Form oform, string fieldname, bool add, bool find, bool update)
        {

            if (add == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (find == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (update)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

        }

        public void Matrix_Addrow(SAPbouiCOM.Matrix omatrix, string colname = "", string rowno_name = "", bool Error_Needed = false)
        {
            try
            {
                bool addrow = false;

                if (omatrix.VisualRowCount == 0)
                {
                    addrow = true;
                    goto addrow;
                }
                if (string.IsNullOrEmpty(colname))
                {
                    addrow = true;
                    goto addrow;
                }
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific).String != "")
                {
                    addrow = true;
                    goto addrow;
                }

                addrow:
                ;

                if (addrow == true)
                {
                    omatrix.AddRow(1);
                    omatrix.ClearRowData(omatrix.VisualRowCount);
                    if (!string.IsNullOrEmpty(rowno_name))
                      ((SAPbouiCOM.EditText) omatrix.Columns.Item("#").Cells.Item(omatrix.VisualRowCount).Specific).String =Convert.ToString(omatrix.VisualRowCount);
                }
                else if (Error_Needed == true)
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception ex)
            {

            }
        }

        public void OpenFile(string ServerPath, string ClientPath)
        {
            try
            {
                System.Diagnostics.Process oProcess = new System.Diagnostics.Process();
                try
                {
                    oProcess.StartInfo.FileName = ServerPath;
                    oProcess.Start();
                }
                catch (Exception ex1)
                {
                    try
                    {
                        oProcess.StartInfo.FileName = ClientPath;
                        oProcess.Start();
                    }
                    catch (Exception ex2)
                    {
                        clsModule.objaddon.objapplication.StatusBar.SetText("" + ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    finally
                    {
                    }
                }
                finally
                {
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }

        public void OpenAttachment(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSAttch, int PvalRow)
        {
            try
            {
                if (PvalRow <= oMatrix.VisualRowCount & PvalRow != 0)
                {
                    int RowIndex = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) - 1;
                    string strServerPath, strClientPath;

                    strServerPath = oDBDSAttch.GetValue("U_TrgtPath", RowIndex) + @"\" + oDBDSAttch.GetValue("U_FileName", RowIndex) + "." + oDBDSAttch.GetValue("U_FileExt", RowIndex);
                    strClientPath = oDBDSAttch.GetValue("U_SrcPath", RowIndex) + @"\" + oDBDSAttch.GetValue("U_FileName", RowIndex) + "." + oDBDSAttch.GetValue("U_FileExt", RowIndex);
                    // Open Attachment File
                    OpenFile(strServerPath, strClientPath);
                }
            }
            catch (Exception ex)
            {
               clsModule.objaddon.objapplication.StatusBar.SetText("OpenAttachment Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }


        public void ShowFolderBrowser()
        {
            System.Diagnostics.Process[] MyProcs;
            OpenFileDialog OpenFile = new OpenFileDialog();
            try
            {
                OpenFile.Multiselect = false;
                OpenFile.Filter = "All files(*.)|*.*"; // "|*.*"
                int filterindex = 0;
                try
                {
                    filterindex = 0;
                }
                catch (Exception ex)
                {
                }
                OpenFile.FilterIndex = filterindex;
                // OpenFile.RestoreDirectory = True
                OpenFile.InitialDirectory = clsModule.objaddon.objcompany.AttachMentPath; // "\\newton.tmicloud.net\DB4SHARE\OEC_TEST\Attachments\"
                MyProcs = Process.GetProcessesByName("SAP Business One");

                // *******  Modified on 09-Mar-2012 By parthiban ********

                // If two or more company opened at the same time,  Dialog is  not opening 
                // Changed Conditon   to >= 1
                // Added Condition --If comname(1).ToString.Trim.ToUpper = com Then -- to open dialog
                // only for this company

                // If MyProcs.Length = 1 Then
                if (MyProcs.Length >= 1)
                {
                    for (int i = 0; i <= MyProcs.Length - 1; i++)
                    {
                        string[] comname = MyProcs[i].MainWindowTitle.ToString().Split(Convert.ToChar(@"-"));
                        if (comname[0] == "")
                            continue;
                        // Open dialog only for the company where the button is clicked
                        string com = clsModule.objaddon.objcompany.CompanyName.ToUpper();
                        if (comname[0].ToString().Trim().ToUpper() == com)
                        {
                            WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                            System.Windows.Forms.DialogResult ret = OpenFile.ShowDialog(MyWindow);
                            if (ret == System.Windows.Forms.DialogResult.OK)
                                BankFileName = OpenFile.FileName;
                            else
                                System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message);
                BankFileName = "";
            }
            finally
            {
                OpenFile.Dispose();
            }
        }

        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private IntPtr _hwnd;

            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }

            public System.IntPtr Handle
            {
                get
                {
                    return _hwnd;
                }
            }
        }

        public string FindFile()
        {
            System.Threading.Thread ShowFolderBrowserThread;
            try
            {
                ShowFolderBrowserThread = new System.Threading.Thread( ShowFolderBrowser);

                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }

                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                    // ShowFolderBrowserThread.Sleep(100)
                    Thread.Sleep(100);
                }

                if (BankFileName != "")
                    return BankFileName;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.MessageBox("File Find  Method Failed : " + ex.Message);
            }
            return "";
        }

        public string GetServerDate()
        {
            try
            {
                SAPbobsCOM.SBObob rsetBob =(SAPbobsCOM.SBObob) clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset rsetServerDate = (SAPbobsCOM.Recordset) clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rsetServerDate = rsetBob.Format_StringToDate(clsModule.objaddon.objapplication.Company.ServerDate);
                DateTime DocDate = Convert.ToDateTime(rsetServerDate.Fields.Item(0).Value);

                return DocDate.ToString("yyyyMMdd");// Convert.ToString(rsetServerDate.Fields.Item(0).Value); //Convert.ToString(rsetServerDate.Fields.Item(0).Value);//.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                return "";
            }
            finally
            {
            }
        }

        public bool SetAttachMentFile(SAPbouiCOM.Form oForm, SAPbouiCOM.DBDataSource oDBDSHeader, SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSAttch)
        {
            try
            {
                if (clsModule.objaddon.objcompany.AttachMentPath.Length <= 0)
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]");
                    return false;
                }

                string strFileName = FindFile();
                if (strFileName.Equals("") == false)
                {
                    string[] FileExist = strFileName.Split(Convert.ToChar(@"\"));
                    string FileDestPath = clsModule.objaddon.objcompany.AttachMentPath + FileExist[FileExist.Length - 1];

                    if (File.Exists(FileDestPath))
                    {
                        long LngRetVal = clsModule.objaddon.objapplication.MessageBox("A file with this name already exists,would you like to replace this?  " + FileDestPath + " will be replaced.", 1, "Yes", "No");
                        if (LngRetVal != 1)
                            return false;
                    }
                    string[] fileNameExt = FileExist[FileExist.Length - 1].Split(Convert.ToChar("."));
                    string ScrPath = clsModule.objaddon.objcompany.AttachMentPath;
                    ScrPath = ScrPath.Substring(0, ScrPath.Length - 1);
                    string TrgtPath = strFileName.Substring(0, strFileName.LastIndexOf(@"\"));

                    oMatrix.AddRow();
                    oMatrix.FlushToDataSource();
                    oDBDSAttch.Offset = oDBDSAttch.Size - 1;
                    oDBDSAttch.SetValue("LineId", oDBDSAttch.Offset,Convert.ToString(oMatrix.VisualRowCount));
                    oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ScrPath);
                    oDBDSAttch.SetValue("U_SrcPath", oDBDSAttch.Offset, TrgtPath);
                    oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt[0]);
                    oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt[1]);
                    oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, GetServerDate());
                    oMatrix.SetLineData(oDBDSAttch.Size);
                    oMatrix.FlushToDataSource();
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                return true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Set AttachMent File Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return false;
            }
            finally
            {
            }
        }

        public void DeleteRowAttachment(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSAttch, int SelectedRowID)
        {
            try
            {
                oDBDSAttch.RemoveRecord(SelectedRowID - 1);
                oMatrix.DeleteRow(SelectedRowID);
                oMatrix.FlushToDataSource();

                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {
                    oMatrix.GetLineData(i);
                    oDBDSAttch.Offset = i - 1;

                    oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset,Convert.ToString(i));
                    oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset,((SAPbouiCOM.EditText) oMatrix.Columns.Item("TrgtPath").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_SrcPath", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Path").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("FileName").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("FileExt").Cells.Item(i).Specific).Value);
                    oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Date").Cells.Item(i).Specific).Value);
                    oMatrix.SetLineData(i);
                    oMatrix.FlushToDataSource();
                }
                // oDBDSAttch.RemoveRecord(oDBDSAttch.Size - 1)
                oMatrix.LoadFromDataSource();

                oForm.Items.Item("btndisp").Enabled = false;
                oForm.Items.Item("btndel").Enabled = false;

                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DeleteRowAttachment Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
            }
        }
        
        public void Load_Combo(string FormUID, SAPbouiCOM.ComboBox comboBox, string Query, string[] Validvalues = null)
        {
            try
            {
                SAPbobsCOM.Recordset objRs;
                string[] split_char;                
                if (comboBox.ValidValues.Count != 0) return;
                
                if (Validvalues.Length > 0)
                {
                    for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                    {
                        if (string.IsNullOrEmpty(Validvalues[i]))
                            continue;
                        split_char = Validvalues[i].Split(Convert.ToChar(","));
                        if (split_char.Length != 2)
                            continue;
                        comboBox.ValidValues.Add(split_char[0], split_char[1]);
                    }
                }

                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRs.DoQuery(Query);
                if (objRs.RecordCount == 0) return;
                for (int i = 0; i < objRs.RecordCount; i++)
                {
                    comboBox.ValidValues.Add(objRs.Fields.Item(0).Value.ToString(), objRs.Fields.Item(1).Value.ToString());
                    objRs.MoveNext();
                }
                comboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public void AddToPermissionTree(string Name, string PermissionID, string FormType, string ParentID, char AddPermission)
        {
            try
            {
                long RetVal;
                string ErrMsg = "";
                SAPbobsCOM.UserPermissionTree oPermission;
                SAPbobsCOM.SBObob objBridge;
                if (ParentID != "")
                {
                    if (clsModule.objaddon.HANA == true)
                        strsql = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as \"Status\" from OUPT Where \"AbsId\"='" + ParentID + "'");
                    else
                        strsql = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as Status from OUPT Where AbsId='" + ParentID + "'");
                    if (strsql == "") return;
                }


                oPermission = (SAPbobsCOM.UserPermissionTree)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);
                objBridge = (SAPbobsCOM.SBObob)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs = objBridge.GetUserList();

                if (oPermission.GetByKey(PermissionID) == false)
                {
                    oPermission.Name = Name;
                    oPermission.PermissionID = PermissionID;
                    oPermission.UserPermissionForms.FormType = FormType;
                    if (ParentID != "") oPermission.ParentID = ParentID;
                    oPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone;
                    RetVal = oPermission.Add();

                    int temp_int = (int)(RetVal);
                    string temp_string = ErrMsg;
                    clsModule.objaddon.objcompany.GetLastError(out temp_int, out temp_string);
                    if (RetVal != 0)
                    {
                        clsModule.objaddon.objapplication.StatusBar.SetText("AddToPermissionTree: " + temp_string, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        //*****************Add Permission To All Active Users*****************
                        if (AddPermission == 'N') return;
                        for (int i = 0; i < objrs.RecordCount ; i++)
                        {
                            //strsql =Convert.ToString(objrs.Fields.Item(0).Value);
                            if (clsModule.objaddon.HANA == true)
                                strsql = "Select \"USERID\" from OUSR Where \"USER_CODE\"='" + Convert.ToString(objrs.Fields.Item(0).Value) + "'";
                            else
                                strsql = "Select USERID from OUSR Where USER_CODE='" + Convert.ToString(objrs.Fields.Item(0).Value) + "'";                            
                            strsql = clsModule.objaddon.objglobalmethods.getSingleValue(strsql);
                            clsModule.objaddon.objglobalmethods.AddPermissionToUsers(Convert.ToInt32(strsql), PermissionID); //clsModule.objaddon.objcompany.UserSignature
                            objrs.MoveNext();
                        }

                    }
                }
                //else
                //{
                //    oPermission.Remove();
                //}
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Permission: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        public void AddPermissionToUsers(int UserCode, string PermissionID)
        {
            try
            {
                SAPbobsCOM.Users oUser = null;
                int lRetCode;
                string sErrMsg = "";

                oUser = ((SAPbobsCOM.Users)(clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)));

                if (oUser.GetByKey(UserCode) == true)
                {
                    oUser.UserPermission.Add();
                    oUser.UserPermission.SetCurrentLine(0);
                    oUser.UserPermission.PermissionID = PermissionID;
                    oUser.UserPermission.Permission = SAPbobsCOM.BoPermission.boper_Full;

                    lRetCode = oUser.Update();

                    clsModule.objaddon.objcompany.GetLastError(out lRetCode, out sErrMsg);
                    if (lRetCode != 0)
                    {
                        clsModule.objaddon.objapplication.StatusBar.SetText("AddPermissionToUser: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }

            }
            catch (Exception ex)
            {

            }
        }

        public void Update_UserFormSettings_UDF(SAPbouiCOM.Form form,string FormID, int UserCode)
        {
            try
            {
                SAPbobsCOM.CompanyService oCmpSrv;
                FormPreferencesService oFormPreferencesService;
                ColumnsPreferences oColsPreferences;
                ColumnsPreferencesParams oColPreferencesParams;
                //get company service
                oCmpSrv = clsModule.objaddon.objcompany.GetCompanyService();
                //get Form Preferences Service
                oFormPreferencesService =(FormPreferencesService) oCmpSrv.GetBusinessService(ServiceTypes.FormPreferencesService);

                //get Columns Preferences Params
                oColPreferencesParams = (ColumnsPreferencesParams)oFormPreferencesService.GetDataInterface(FormPreferencesServiceDataInterfaces.fpsdiColumnsPreferencesParams);

                //set the form id (e.g. A/R invoice=133)
                oColPreferencesParams.FormID = FormID;// "133";

                //set the user id (e.g manager= 1)
                oColPreferencesParams.User = UserCode;// 1;

                //get the Columns Preferences according to the formId & user id
                oColsPreferences = oFormPreferencesService.GetColumnsPreferences(oColPreferencesParams);

                //change the width of all the visible items
                //for (int i = 0; i < form.Items.Count - 1; i++)
                //{
                //    form.Items.Item(i).Visible = false;
                //    form.Items.Item(i).Enabled = false;
                //    strsql= form.Items.Item(i).UniqueID;
                //    break;
                //}
                for (int i = 0; i < oColsPreferences.Count-1; i++)
                {
                    
                    if (oColsPreferences.Item(i).VisibleInForm == BoYesNoEnum.tYES)
                    {
                        oColsPreferences.Item(i).EditableInForm = BoYesNoEnum.tNO;
                        oColsPreferences.Item(i).VisibleInForm = BoYesNoEnum.tNO;
                    }

                }

                //update all changes
                oFormPreferencesService.UpdateColumnsPreferences(oColPreferencesParams, oColsPreferences);
            }
            catch (Exception)
            {

                throw;
            }
        }



    }
}
