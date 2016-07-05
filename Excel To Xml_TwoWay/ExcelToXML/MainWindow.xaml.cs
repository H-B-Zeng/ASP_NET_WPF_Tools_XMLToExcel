﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Collections;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelToXML
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        //選擇Excel file
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "(.xlsx)|*.xlsx";
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {

                txtFIle.Text = dlg.FileName;


                List sheetNameList = new List();

                System.Data.OleDb.OleDbConnection conn = GetConnect(txtFIle.Text);
                conn.Open();

                DataTable sheetList = conn.GetSchema("Tables");

                cmbSheet.DisplayMemberPath = "TABLE_NAME";
                cmbSheet.SelectedValuePath = "TABLE_NAME";
                cmbSheet.ItemsSource = sheetList.DefaultView;
            }
        }
        //產生XML文件
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string tname = cmbSheet.SelectedValue.ToString().Replace("'", string.Empty);

            DataTable line = GetExcelDataTable(txtFIle.Text, tname.Replace("$", string.Empty));

            //DataTable line = GetExcelDataTable(txtFIle.Text, "20150423");

            //寫入XML
            string StartWord = string.Format("<ResourceDictionary xmlns='{0}' xmlns:x='{1}' xmlns:sys='{2}'>", "http://schemas.microsoft.com/winfx/2006/xaml/presentation", "http://schemas.microsoft.com/winfx/2006/xaml", "clr-namespace:System;assembly=mscorlib"); //開始字串
            string ContentWord = "<sys:String x:Key='{0}'>{1}</sys:String>";
            string EndWord = "</ResourceDictionary>";    //結束字串

            string dir = "C:\\res";

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            foreach (DataColumn cols in line.Columns)
            {
                if (cols.ColumnName != "var")
                {
                    FileStream fs = new FileStream(dir + "\\" + cols.ColumnName + ".xaml", FileMode.OpenOrCreate);
                    StreamWriter swWriter = new StreamWriter(fs);
                    //寫入數據
                    swWriter.WriteLine(StartWord);

                    for (int i = 0; i < line.Rows.Count; i++)
                    {
                        string ContentLine = string.Format(ContentWord, line.Rows[i]["var"].ToString(), line.Rows[i][cols.ColumnName].ToString());
                        swWriter.WriteLine(ContentLine);
                    }
                    swWriter.WriteLine(EndWord);
                    swWriter.Close();

                }

            }

        }

        //public DataRow GetExcelDataSheet()
        //{
        /// load excel sheet as DataRow
        //DataRow[] sheetList =
        //    GetConnect().GetSchema("Tables").Select();

        //}

        public System.Data.OleDb.OleDbConnection GetConnect(string path)
        {
            /*Office 2007*/
            string ace = "Microsoft.ACE.OLEDB.12.0";
            /*Office 97 - 2003*/
            string jet = "Microsoft.Jet.OLEDB.4.0";
            string xl2007 = "Excel 12.0 Xml";
            string xl2003 = "Excel 8.0";
            string imex = "IMEX=1";
            /* csv */
            string text = "text";
            string fmt = "FMT=Delimited";
            string hdr = "Yes";
            string conn = "Provider={0};Data Source={1};Extended Properties=\"{2};HDR={3};{4}\";";
            //string select = string.Format("SELECT * FROM [{0}$]", tname);
            //string select = sql;
            string ext = System.IO.Path.GetExtension(path);
            System.Data.OleDb.OleDbDataAdapter oda;
            DataTable dt = new DataTable("data");
            switch (ext.ToLower())
            {
                case ".xlsx":
                    conn = String.Format(conn, ace, System.IO.Path.GetFullPath(path), xl2007, hdr, imex);
                    break;
                case ".xls":
                    conn = String.Format(conn, jet, System.IO.Path.GetFullPath(path), xl2003, hdr, imex);
                    break;
                case ".csv":
                    conn = String.Format(conn, jet, System.IO.Path.GetDirectoryName(path), text, hdr, fmt);
                    //sheet = Path.GetFileName(path);
                    break;
                case ".xaml":
                    conn = String.Format(conn, jet, System.IO.Path.GetDirectoryName(path), text, hdr, fmt);
                    //sheet = Path.GetFileName(path);
                    break;
                default:
                    throw new Exception("File Not Supported!");
            }
            System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(conn);

            return con;
        }

        /// <summary>
        /// 根據Xaml路徑和file名稱，返回xaml的DataTable
        /// </summary>
        /// <param name="path">Excel 檔路徑</param>
        /// <param name="tname">Excel WorkSheet name</param>
        /// <returns>Data Table</returns>
        public DataTable GetExcelDataTable(string path, string tname)
        {
            /*Office 2007*/
            string ace = "Microsoft.ACE.OLEDB.12.0";
            /*Office 97 - 2003*/
            string jet = "Microsoft.Jet.OLEDB.4.0";
            string xl2007 = "Excel 12.0 Xml";
            string xl2003 = "Excel 8.0";
            string imex = "IMEX=1";
            /* csv */
            string text = "text";
            string fmt = "FMT=Delimited";
            string hdr = "Yes";
            string conn = "Provider={0};Data Source={1};Extended Properties=\"{2};HDR={3};{4}\";";
            string select = string.Format("SELECT * FROM [{0}$]", tname);
            //string select = sql;
            string ext = System.IO.Path.GetExtension(path);
            System.Data.OleDb.OleDbDataAdapter oda;
            DataTable dt = new DataTable("data");
            switch (ext.ToLower())
            {
                case ".xlsx":
                    conn = String.Format(conn, ace, System.IO.Path.GetFullPath(path), xl2007, hdr, imex);
                    break;
                case ".xls":
                    conn = String.Format(conn, jet, System.IO.Path.GetFullPath(path), xl2003, hdr, imex);
                    break;
                case ".csv":
                    conn = String.Format(conn, jet, System.IO.Path.GetDirectoryName(path), text, hdr, fmt);
                    //sheet = Path.GetFileName(path);
                    break;
                default:
                    throw new Exception("File Not Supported!");
            }
            System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(conn);
            con.Open();
            //select = string.Format(select, sql);
            oda = new System.Data.OleDb.OleDbDataAdapter(select, con);
            oda.Fill(dt);
            con.Close();
            return dt;
        }

        #region XML To Excel

        /// <summary>
        /// select xaml files,show to View ListBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectXamlFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog chrooseFileDialog = new OpenFileDialog();
            chrooseFileDialog.DefaultExt = ".xaml";
            chrooseFileDialog.Filter = "XML files(.xaml; .xml)|*.xaml; *.xml";
            chrooseFileDialog.Multiselect = true;
            chrooseFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Nullable<bool> result = chrooseFileDialog.ShowDialog();
            string defaultSaveExcelPath = string.Empty;

            String errorMsg = "";

            try
            {
                if (result == true)
                {
                    List<string> xamlNameList = chrooseFileDialog.FileNames.ToList();

                    if (xamlNameList != null)
                    {
                        lbFileList.Items.Clear();
                        foreach (string fileName in xamlNameList)
                        {
                            lbFileList.Items.Add(fileName);

                            if (string.IsNullOrEmpty(txtExcelSavePath.Text))
                            {
                                txtExcelSavePath.Text = fileName.Replace("xaml", "xlsx").Replace("xml", "xlsx");
                            }
                        }

                        errorMsg = VerificationXMLFormat(xamlNameList);
                        if (!string.IsNullOrEmpty(errorMsg))
                        {
                            MessageBox.Show("please,Check Xml format \r\n\r\n File Name:" + errorMsg, "XML Info");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// 將多個xml轉成單一Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnXamlToExcel(object sender, RoutedEventArgs e)
        {
            //contain path and Name
            string getXmlFullPath;
            string getXmlName;
            string errorMsg = string.Empty;
            bool exportSwitch = true;
            DataTable FirstDt = new DataTable();
            DataTable TempDt = new DataTable();

            try
            {
                //還沒選檔案
                if (lbFileList.Items.Count == 0)
                {
                    exportSwitch = false;
                    MessageBox.Show("Please Choose XML File", "XML Info");
                }
                else if (string.IsNullOrEmpty(txtExcelSavePath.Text)) //還未設定存檔路徑
                {
                    exportSwitch = false;
                    MessageBox.Show("Please check export excel path", "Info");
                }

                if (lbFileList.Items.Count > 0 && exportSwitch)
                {
                    errorMsg = VerificationXMLFormat(lbFileList.Items.Cast<String>().ToList());
                    if (string.IsNullOrEmpty(errorMsg))
                    {
                        for (int i = 0; i < lbFileList.Items.Count; i++)
                        {
                            getXmlFullPath = lbFileList.Items[i].ToString();
                            getXmlName = System.IO.Path.GetFileNameWithoutExtension(getXmlFullPath);
                            if (!string.IsNullOrEmpty(getXmlFullPath))
                            {
                                FirstDt = XmlToDataTable(getXmlFullPath);
                                if (FirstDt.Columns.Contains("String_Text"))
                                {
                                    FirstDt.Columns["String_Text"].ColumnName = getXmlName;
                                }
                                FirstDt.PrimaryKey = new DataColumn[] { FirstDt.Columns[0] };
                                FirstDt.AcceptChanges();
                                UpdateTableColumns(TempDt, FirstDt);
                                TempDt.Merge(FirstDt, false, MissingSchemaAction.Add);
                            }
                        }

                        //Export Excel File
                        resultMsg result = DataTableToExcelFile(TempDt);
                        if (result.isSuccess)
                        {
                            MessageBox.Show(result.Msg, "Info");
                        }
                        else
                        {
                            MessageBox.Show(result.Msg, "error");
                        }

                    }
                    else
                    {
                        MessageBox.Show("please,Check Xml format \r\n\r\n File Name:" + errorMsg, "XML Info");
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                throw;
            }
        }

        /// <summary>
        /// Before xml to table,check xml format
        /// </summary>
        /// <param name="XmlfileNameList">all file name</param>
        /// <returns>xmlFormatErrorMsg</returns>
        public string VerificationXMLFormat(List<string> XmlfileNameList)
        {
            string xmlFormatErrorMsg = string.Empty;
            foreach (string fileName in XmlfileNameList)
            {
                FileInfo getFileInfo = new FileInfo(fileName);
                if (getFileInfo.Length < 16)
                {
                    xmlFormatErrorMsg += getFileInfo.Name + ",";
                }
            }

            return xmlFormatErrorMsg;
        }

        /// <summary>
        /// Xml Convert to DataTable
        /// </summary>
        /// <param name="XmlFullName">Contain path and name</param>
        /// <returns>DataTable</returns>
        public DataTable XmlToDataTable(string XmlFullName)
        {
            DataTable dt;
            DataSet ds = new DataSet();
            try
            {
                FileStream fsReadXml = new FileStream(XmlFullName, FileMode.Open);
                XmlTextReader xmlReader = new XmlTextReader(fsReadXml);
                ds.ReadXml(xmlReader);
                xmlReader.Close();
                fsReadXml.Close();
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine("DataTable XmlToDataTable error=" + ex.ToString());
                throw ex;
            }
            return dt;
        }

        /// <summary>
        /// SecondaryTable's Columns Add to MasterTable Columns
        /// </summary>
        /// <param name="MasterTable">Update MasterTable</param>
        /// <param name="MasterTable">MasterTable</param>
        /// <returns>MasterTable</returns>
        public DataTable UpdateTableColumns(DataTable MasterTable, DataTable SecondaryTable)
        {
            DataColumnCollection MasterTableColumns = MasterTable.Columns;
            DataColumnCollection SecondaryTableColumns = SecondaryTable.Columns;
            DataColumn dc;
            try
            {
                foreach (var ColumnName in SecondaryTableColumns)
                {
                    if (!MasterTableColumns.Contains(ColumnName.ToString()))
                    {
                        dc = new DataColumn();
                        dc.ColumnName = ColumnName.ToString();
                        MasterTable.Columns.Add(dc);
                    }
                }
                MasterTable.AcceptChanges();
            }
            catch (Exception)
            {

                throw;
            }
            return MasterTable;
        }

        /// <summary>
        ///  DataTable To ExcelFile,Will Save File
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="filePath">Save Paht</param>
        private resultMsg DataTableToExcelFile(DataTable dt)
        {

            bool SavetSwitch = false;
            string getExcelSaveFullPath = txtExcelSavePath.Text;
            string getExcelSavePath = string.Empty;
            resultMsg result = new resultMsg();

            int ColumnWidth = 35;

            try
            {
                FileInfo filePath = new FileInfo(getExcelSaveFullPath);

                if (!string.IsNullOrEmpty(getExcelSaveFullPath))
                {
                    getExcelSavePath = System.IO.Path.GetDirectoryName(txtExcelSavePath.Text);
                    if (!Directory.Exists(getExcelSavePath))
                    {
                        Directory.CreateDirectory(getExcelSavePath);
                    }

                    //if CreateDirectory fail
                    if (!Directory.Exists(getExcelSavePath))
                    {
                        getExcelSavePath = string.Empty;
                        result.isSuccess = false;
                        result.Msg = "Please check excel save path";
                    }
                    else
                    {
                        SavetSwitch = true;
                    }
                }

                if (!File.Exists(getExcelSaveFullPath) && SavetSwitch)
                {
                    ExcelPackage ep = new ExcelPackage(filePath);
                    ExcelWorksheet ws;

                    if (dt.TableName != string.Empty)
                    {
                        ws = ep.Workbook.Worksheets.Add(dt.TableName);
                    }
                    else
                    {
                        ws = ep.Workbook.Worksheets.Add("Sheet1");
                    }

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = dt.Columns[i].ColumnName;
                        ws.Column(i + 1).Width = ColumnWidth;
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            ws.Cells[i + 2, j + 1].Value = dt.Rows[i][j].ToString();
                        }
                    }

                    ws.Cells[1, 1, dt.Rows.Count + 2, dt.Columns.Count + 1].Style.Font.Size = 12;
                    ws.Cells[1, 1, 2, dt.Columns.Count + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[3, 1, dt.Rows.Count + 2, dt.Columns.Count + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[1, 1, dt.Rows.Count + 2, dt.Columns.Count + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells.Style.WrapText = true;
                    ep.Save();
                    ep.Dispose();
                    ep = null;
                    result.isSuccess = true;
                    result.Msg = "Export excel Success";

                }
                else
                {
                    result.isSuccess = false;
                    result.Msg = "File already exists : " + getExcelSaveFullPath;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return result;
        }

        public class resultMsg
        {
            public bool isSuccess { get; set; }
            public string Msg { get; set; }
        }

        #endregion

    }
}