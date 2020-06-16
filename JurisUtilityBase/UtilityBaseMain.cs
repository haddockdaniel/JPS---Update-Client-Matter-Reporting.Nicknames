using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;
using JurisSVR;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public string fileName = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            if (string.IsNullOrEmpty(fileName))
                MessageBox.Show("Please select an Excel file before proceeding", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                CliObj cli = null;
                List<CliObj> clients = new List<CliObj>();

                int lastUsedRow = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                int count = 1;

                for (int i = 3; i <= lastUsedRow; i++)
                {
                    bool cliCodeExists = true;
                    for (int j = 2; j <= 12; j++)
                    {
                        switch (j)
                        {
                            case 2:
                                cli = new CliObj();
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    cli.clicode = xlRange.Cells[i, j].Value2.ToString();
                                else
                                    cliCodeExists = false;
                                break;
                            case 3:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.newName = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.newName = cli.newName.Replace("\r", " ");
                                    cli.newName = cli.newName.Replace("\n", " ");
                                    cli.newName = cli.newName.Replace("\"", " ");
                                    cli.newName = cli.newName.Replace("%", " ");
                                    cli.newName = cli.newName.Replace("#", " ");
                                    cli.newName = cli.newName.Replace("*", " ");
                                    cli.newName = cli.newName.Replace("$", " ");
                                    cli.newName = cli.newName.Replace("@", " ");
                                    cli.updateCliName = true;
                                }
                                else
                                {
                                    cli.updateCliName = false;
                                }
                                break;
                            case 4:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.BF01 = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.BF01 = cli.BF01.Replace("\r", " ");
                                    cli.BF01 = cli.BF01.Replace("\n", " ");
                                    cli.BF01 = cli.BF01.Replace("\"", " ");
                                    cli.BF01 = cli.BF01.Replace("%", " ");
                                    cli.BF01 = cli.BF01.Replace("#", " ");
                                    cli.BF01 = cli.BF01.Replace("*", " ");
                                    cli.BF01 = cli.BF01.Replace("$", " ");
                                    cli.BF01 = cli.BF01.Replace("@", " ");
                                }
                                else
                                {
                                    cli.BF01 = "";
                                    cli.flag1 = "ZZ3S5Vb";
                                }
                                break;
                            case 5:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.BF03 = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.BF03 = cli.BF03.Replace("\r", " ");
                                    cli.BF03 = cli.BF03.Replace("\n", " ");
                                    cli.BF03 = cli.BF03.Replace("\"", " ");
                                    cli.BF03 = cli.BF03.Replace("%", " ");
                                    cli.BF03 = cli.BF03.Replace("#", " ");
                                    cli.BF03 = cli.BF03.Replace("*", " ");
                                    cli.BF03 = cli.BF03.Replace("$", " ");
                                    cli.BF03 = cli.BF03.Replace("@", " ");
                                }
                                else
                                {
                                    cli.BF03 = "";
                                    cli.flag3 = "ZZ3S5Vb";
                                }
                                break;
                            case 6:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.BF04 = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.BF04 = cli.BF04.Replace("\r", " ");
                                    cli.BF04 = cli.BF04.Replace("\n", " ");
                                    cli.BF04 = cli.BF04.Replace("\"", " ");
                                    cli.BF04 = cli.BF04.Replace("%", " ");
                                    cli.BF04 = cli.BF04.Replace("#", " ");
                                    cli.BF04 = cli.BF04.Replace("*", " ");
                                    cli.BF04 = cli.BF04.Replace("$", " ");
                                    cli.BF04 = cli.BF04.Replace("@", " ");
                                }
                                else
                                {
                                    cli.BF04 = "";
                                    cli.flag4 = "ZZ3S5Vb";
                                }
                                break;
                            case 7:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.BF05 = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.BF05 = cli.BF05.Replace("\r", " ");
                                    cli.BF05 = cli.BF05.Replace("\n", " ");
                                    cli.BF05 = cli.BF05.Replace("\"", " ");
                                    cli.BF05 = cli.BF05.Replace("%", " ");
                                    cli.BF05 = cli.BF05.Replace("#", " ");
                                    cli.BF05 = cli.BF05.Replace("*", " ");
                                    cli.BF05 = cli.BF05.Replace("$", " ");
                                    cli.BF05 = cli.BF05.Replace("@", " ");
                                }
                                else
                                {
                                    cli.BF05 = "";
                                    cli.flag5 = "ZZ3S5Vb";
                                }
                                break;
                            case 8:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.BF06 = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.BF06 = cli.BF06.Replace("\r", " ");
                                    cli.BF06 = cli.BF06.Replace("\n", " ");
                                    cli.BF06 = cli.BF06.Replace("\"", " ");
                                    cli.BF06 = cli.BF06.Replace("%", " ");
                                    cli.BF06 = cli.BF06.Replace("#", " ");
                                    cli.BF06 = cli.BF06.Replace("*", " ");
                                    cli.BF06 = cli.BF06.Replace("$", " ");
                                    cli.BF06 = cli.BF06.Replace("@", " ");
                                }
                                else
                                {
                                    cli.BF06 = "";
                                    cli.flag6 = "ZZ3S5Vb";
                                }
                                break;
                            case 9:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    cli.BF07 = xlRange.Cells[i, j].Value2.ToString().Replace("'", " ");
                                    cli.BF07 = cli.BF07.Replace("\r", " ");
                                    cli.BF07 = cli.BF07.Replace("\n", " ");
                                    cli.BF07 = cli.BF07.Replace("\"", " ");
                                    cli.BF07 = cli.BF07.Replace("%", " ");
                                    cli.BF07 = cli.BF07.Replace("#", " ");
                                    cli.BF07 = cli.BF07.Replace("*", " ");
                                    cli.BF07 = cli.BF07.Replace("$", " ");
                                    cli.BF07 = cli.BF07.Replace("@", " ");
                                }
                                else
                                {
                                    cli.BF07 = "";
                                    cli.flag7 = "ZZ3S5Vb";
                                }
                                break;
                            case 10:
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && !String.IsNullOrEmpty(xlRange.Cells[i, j].Value2.ToString()))
                                {
                                    try
                                    {
                                        int mat = Int32.Parse(xlRange.Cells[i, j].Value2.ToString());
                                        if (mat == 0)
                                            cli.changeMatter = true;
                                        else
                                            cli.changeMatter = false;
                                    }
                                    catch (Exception ex1)
                                    {
                                        cli.changeMatter = false;
                                    }
                                }
                                else
                                    cli.changeMatter = false;
                                break;

                        }
                        //write the value to the console
                        if (!cliCodeExists)
                            break;

                    }
                    if (cliCodeExists)
                        clients.Add(cli);
                    UpdateStatus("Accessing Spreadsheet", count, lastUsedRow * 2);
                    count++;
                }

                //close and release
                xlWorkbook.Close();

                //quit and release
                xlApp.Quit();

                UpdateStatus("Updating Database", count, lastUsedRow * 2);

                foreach (CliObj cl in clients)
                {
                    string sql = "";

                    if (cl.updateCliName && !checkBox1.Checked)
                        sql = "update client set CliNickName = left('" + cl.newName + "', 30), CliReportingName = left('" + cl.newName + "', 30), CliBillingField01 = case when '" + cl.flag1 + "' = 'ZZ3S5Vb' then CliBillingField01 else '" + cl.BF01 + "' end, " +
                            " CliBillingField03 = case when '" + cl.flag3 + "' = 'ZZ3S5Vb' then CliBillingField03 else '" + cl.BF03 + "' end,CliBillingField04 = case when '" + cl.flag4 + "' = 'ZZ3S5Vb' then CliBillingField04 else '" + cl.BF04 + "' end, " +
                            " CliBillingField05 = case when '" + cl.flag5 + "' = 'ZZ3S5Vb' then CliBillingField05 else '" + cl.BF05 + "' end, " +
                            " CliBillingField06 = case when '" + cl.flag6 + "' = 'ZZ3S5Vb' then CliBillingField06 else '" + cl.BF06 + "' end,CliBillingField07 = case when '" + cl.flag7 + "' = 'ZZ3S5Vb' then CliBillingField07 else '" + cl.BF07 + "' end " + 
                            " where cast(dbo.jfn_FormatClientCode(clicode) as varchar(8)) = '" + cl.clicode + "'";
                    else if (cl.updateCliName && checkBox1.Checked)
                        sql = "update client set CliNickName = left('" + cl.newName + "', 30), CliReportingName = left('" + cl.newName + "', 30), CliBillingField01 = '" + cl.BF01 + "', " +
                            " CliBillingField03 = '" + cl.BF03 + "',CliBillingField04 = '" + cl.BF04 + "',CliBillingField05 = '" + cl.BF05 + "', " +
                            " CliBillingField06 = '" + cl.BF06 + "',CliBillingField07 = '" + cl.BF07 + "' " +
                            " where cast(dbo.jfn_FormatClientCode(clicode) as varchar(8)) = '" + cl.clicode + "'";
                    else if (!cl.updateCliName && !checkBox1.Checked)
                        sql = "update client set CliBillingField01 = case when '" + cl.flag1 + "' = 'ZZ3S5Vb' then CliBillingField01 else '" + cl.BF01 + "' end, " +
                            " CliBillingField03 = case when '" + cl.flag3 + "' = 'ZZ3S5Vb' then CliBillingField03 else '" + cl.BF03 + "' end,CliBillingField04 = case when '" + cl.flag4 + "' = 'ZZ3S5Vb' then CliBillingField04 else '" + cl.BF04 + "' end, " +
                            " CliBillingField05 = case when '" + cl.flag5 + "' = 'ZZ3S5Vb' then CliBillingField05 else '" + cl.BF05 + "' end, " +
                            " CliBillingField06 = case when '" + cl.flag6 + "' = 'ZZ3S5Vb' then CliBillingField06 else '" + cl.BF06 + "' end,CliBillingField07 = case when '" + cl.flag7 + "' = 'ZZ3S5Vb' then CliBillingField07 else '" + cl.BF07 + "' end " +
                            " where cast(dbo.jfn_FormatClientCode(clicode) as varchar(8)) = '" + cl.clicode + "'";
                    else if (!cl.updateCliName && checkBox1.Checked)
                        sql = "update client set CliBillingField01 = '" + cl.BF01 + "', " +
                            " CliBillingField03 = '" + cl.BF03 + "',CliBillingField04 = '" + cl.BF04 + "',CliBillingField05 = '" + cl.BF05 + "', " +
                            " CliBillingField06 = '" + cl.BF06 + "',CliBillingField07 = '" + cl.BF07 + "' " +
                    " where cast(dbo.jfn_FormatClientCode(clicode) as varchar(8)) = '" + cl.clicode + "'";
                    _jurisUtility.ExecuteNonQuery(0, sql);

                    if (cl.updateCliName && cl.changeMatter)
                    {
                        sql = "update matter set MatNickName = left('" + cl.newName + "', 30), MatReportingName = left('" + cl.newName + "', 30) " +
                            " where matclinbr in (select clisysnbr from client where cast(dbo.jfn_FormatClientCode(clicode) as varchar(8)) = '" + cl.clicode + "') and matcode = '000000000000'";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                    }
                    UpdateStatus("Updating Database", count, lastUsedRow * 2);
                    count++;
                }

                UpdateStatus("Updating Database", lastUsedRow * 2, lastUsedRow * 2);
                count++;



                UpdateStatus("All fields updated.", 1, 1);

                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);

                clients.Clear();
                fileName = "";
            }
        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Browse for Excel File";
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }
        }
    }
}
