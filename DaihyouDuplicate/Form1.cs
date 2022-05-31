using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using MySql.Data.MySqlClient;

namespace DaihyouDuplicate
{
    public partial class Form1 : Form
    {
        private MySqlConnection connection;
        string MyCon;
        private string path = "";
        private XmlElement root;
        private string serverName = "";
        private string dbName = " ";
        private string dbUser = " ";
        private string dbPass = " ";
        public Form1()
        {
            InitializeComponent();
        }
        public void Connection()//get connection
        {
            try
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                path = path + "\\DEMO20" + "\\demo20.xml";
                if (File.Exists(path))
                {
                    //fileNotExists = false;
                    XmlDocument doc = new XmlDocument();
                    doc.Load(path);
                    root = doc.DocumentElement;
                    serverName = root.GetElementsByTagName("DB-SERV")[0].InnerText;
                    dbName = root.GetElementsByTagName("DB-NAME")[0].InnerText;
                    dbUser = root.GetElementsByTagName("DB-USER")[0].InnerText;
                    dbPass = root.GetElementsByTagName("DB-PASS")[0].InnerText;
                    if (dbPass == "")
                    {
                        dbPass = "hf71v6n2";
                    }
                    MyCon = "server=" + serverName + "; user=" + dbUser + " ; password=" + dbPass + "; database=" + dbName + ";";
                    connection = new MySqlConnection(MyCon);
                    connection.Open();
                    connection.Close();
                }
                else
                {
                    MessageBox.Show("データベースに接続できません、管理者にお問合せ下さい。", "確認");
                    //fileNotExists = true;
                    return;
                }
            }
            catch
            {
                MessageBox.Show("データベースに接続できません、管理者にお問合せ下さい。", "確認");
                //fileNotExists = true;
                return;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            this.Hide();
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet1;
            Microsoft.Office.Interop.Excel.Worksheet oSheet2;
            Microsoft.Office.Interop.Excel.Range oRng1;
            Microsoft.Office.Interop.Excel.Range oRng2;
            Object[,] excelArray1;
            Object[,] excelArray2;
            object missing = Type.Missing;
            Connection();

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                //Get a new workbook.
                oWB = oXL.Workbooks.Add(missing);

                //string sql1 = " SELECT cBUKKEN,nJUNBAN ,COUNT(*)" +
                //        " FROM m_bu_file" +
                //        " GROUP BY cBUKKEN,nJUNBAN" +
                //        " HAVING COUNT(*) > 1; ";
                string sql1 = " SELECT cMITUMORI,nJUNBAN ,COUNT(*)" +
                       " FROM m_mitsu_file" +
                       " GROUP BY cMITUMORI,nJUNBAN" +
                       " HAVING COUNT(*) > 1; ";
                MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, connection);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);
                if (dt1.Rows.Count != 0)
                {
                    excelArray1 = new Object[dt1.Rows.Count + 1, dt1.Columns.Count - 1];
                    excelArray1[0, 0] = "cMITUMORI";
                    excelArray1[0, 1] = "nJUNBAN";
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        excelArray1[i + 1, 0] = "'" + dt1.Rows[i][0].ToString();
                        excelArray1[i + 1, 1] = dt1.Rows[i][1].ToString();
                    }
                    //oSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Worksheets[1];
                    oSheet1 = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    //oSheet1 = oWB.Sheets.Add(missing, missing, 1, missing)
                    //  as Microsoft.Office.Interop.Excel.Worksheet;
                    //string date = DateTime.Now.ToString("yyyy/MM/dd") + "_DaihyouDuplicate";
                    oSheet1.Name = "r_mitsu_file";//r_bu_file

                    Microsoft.Office.Interop.Excel.Range t1 = (Microsoft.Office.Interop.Excel.Range)oSheet1.Cells[1, 1];
                    Microsoft.Office.Interop.Excel.Range t2 = (Microsoft.Office.Interop.Excel.Range)oSheet1.Cells[dt1.Rows.Count + 1, 2];
                    oRng1 = oSheet1.get_Range(t1, t2);
                    oRng1.Value = excelArray1;
                    oRng1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    // oRng1.Borders.LineStyle = DataGridLineStyle.Solid;
                    oSheet1.Cells.EntireColumn.AutoFit();
                }
                string sql2 = " SELECT cBUKKEN,nJUNBAN ,COUNT(*)" +
                       " FROM m_bu_file" +
                       " GROUP BY cBUKKEN,nJUNBAN" +
                       " HAVING COUNT(*) > 1; ";
                //string sql2 = " SELECT cMITUMORI,nJUNBAN ,COUNT(*)" +
                //       " FROM m_mitsu_file" +
                //       " GROUP BY cMITUMORI,nJUNBAN" +
                //       " HAVING COUNT(*) > 1; ";
                MySqlDataAdapter da2 = new MySqlDataAdapter(sql2, connection);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                if (dt2.Rows.Count != 0)
                {
                    excelArray2 = new Object[dt2.Rows.Count + 1, dt2.Columns.Count - 1];
                    excelArray2[0, 0] = "cBUKKEN";
                    excelArray2[0, 1] = "nJUNBAN";
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        excelArray2[i + 1, 0] = "'" + dt2.Rows[i][0].ToString();
                        excelArray2[i + 1, 1] = dt2.Rows[i][1].ToString();
                    }

                    //oSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Worksheets[2];
                    if (dt1.Rows.Count != 0)
                    {
                        oSheet2 = oWB.Sheets.Add(missing, missing, 1, missing)
                      as Microsoft.Office.Interop.Excel.Worksheet;
                    }
                    else
                    {
                        oSheet2 = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    }
                    oSheet2.Name = "r_bu_file";//r_mitsu_file

                    Microsoft.Office.Interop.Excel.Range t11 = (Microsoft.Office.Interop.Excel.Range)oSheet2.Cells[1, 1];
                    Microsoft.Office.Interop.Excel.Range t22 = (Microsoft.Office.Interop.Excel.Range)oSheet2.Cells[dt2.Rows.Count + 1, 2];
                    oRng2 = oSheet2.get_Range(t11, t22);
                    oRng2.Value = excelArray2;
                    oRng2.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //oRng2.Borders.LineStyle = DataGridLineStyle.Solid;
                    oSheet2.Cells.EntireColumn.AutoFit();
                }
                if (dt1.Rows.Count != 0 || dt2.Rows.Count != 0)
                {
                    string executing = Path.GetDirectoryName(Application.ExecutablePath);
                   
                   string Excelname = DateTime.Now.ToString("yyyyMMdd") + "_DaihyouDuplicate_" + dbName + ".xlsx";
                    ////oWB.SaveAs(dbName + ".xlsx",
                    //// Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, "",
                    ////"", false, false,
                    ////Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    ////true, Type.Missing, Type.Missing, Type.Missing);

                    // oWB.SaveAs(executing, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
                    //oXL.Visible = true;
                    //oXL.UserControl = true;
                    //oWB.Close(true, missing, missing);
                    //oXL.Quit();
                    // wbook.SaveAs("c:\\file_path\\file_name", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                    SaveFileDialog saveDialog = new SaveFileDialog();
                    object misValue = System.Reflection.Missing.Value;
                    saveDialog.InitialDirectory = executing;
                    saveDialog.FileName = Excelname;
                    saveDialog.Filter = "Excel|*.xlsx";
                    saveDialog.FilterIndex = 0;
                    saveDialog.RestoreDirectory = true;
                    // excel.Visible = true;
                    if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) //dialogbox OK click
                    {

                        try//save excel
                        {
                            oWB.SaveAs(saveDialog.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                              false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                              Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            oWB.Close(true, misValue, misValue);
                            oXL.Quit();
                           // MessageBox.Show("出力しました。", "確認", MessageBoxButtons.OK);
                        }
                        catch
                        {
                        }
                    }
                    else//dialogbox CANCEL click
                    {
                        object missvalue = System.Reflection.Missing.Value;
                        oWB.Close(false, missvalue, missvalue);
                        oXL.Quit();
                    }




                }
            }
            catch (Exception theException)
            {

            }
            this.Close();
        }
    }
}
