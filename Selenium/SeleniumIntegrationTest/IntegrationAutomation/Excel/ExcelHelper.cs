using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace IntegrationAutomation.Excel
{
    public class ExcelHelper : OperationInterface
    {
        ILog _log { get; set; }
        public ExcelHelper(ILog log)
        {
            _log = log;
        }

        public void PerforAction(ActionRowEntity action)
        {
            _log.Log("Before Action: " + action.ToString());
            try
            {
                var commandType = action.CommandType;
                switch (action.CommandName)
                {
                    case "Compare File":
                        if(commandType == "Last Modified File")
                        {
                            string originalFile = action.CommandValue;
                            string compareFile = GetLastModifiedFile(action.CommandComponent);
                            bool result = CompareFile(originalFile, compareFile);
                            if(!result)
                            {
                                throw new ArgumentException("Two files are different:\r\n" + originalFile + "\r\n" + compareFile);
                            }
                        }
                        break;
                    case "Compare File With Rearrange":
                        if(commandType == "Last Modified File")
                        {
                            string originalFile = action.CommandValue;
                            string compareFile = GetLastModifiedFile(action.CommandComponent);
                            bool result = CompareFile(originalFile, compareFile, true);
                            if(!result)
                            {
                                throw new ArgumentException("Two files are different:\r\n" + originalFile + "\r\n" + compareFile);
                            }
                        }
                        break;
                    default:
                        throw new ArgumentException("Command Name is not implemented");
                }
            }
            catch (Exception ex)
            {
                _log.Log("Exception:" + ex.ToString());
                throw new Exception(action.ToString() + "\r\n" + ex.ToString());
            }

            WaitTime(action.WaitTime);
        }

        public void WaitTime(string time)
        {
            if (!string.IsNullOrEmpty(time))
            {
                if (Int32.TryParse(time, out int seconds))
                {
                    _log.Log("Wait Seconds:" + seconds.ToString());
                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(seconds);
                    Thread.Sleep(seconds * 1000);
                }
                else
                {
                    throw new ArgumentException("Wait time is not an integer");
                }
            }
        }

        public bool CompareDataTable(DataTable originalDataTable, DataTable compareDataTable)
        {
            if (originalDataTable.Rows.Count != compareDataTable.Rows.Count || originalDataTable.Columns.Count != compareDataTable.Columns.Count)
                return false;

            for (int i = 0; i < originalDataTable.Rows.Count; i++)
            {
                for (int c = 0; c < originalDataTable.Columns.Count; c++)
                {
                    if (!Equals(originalDataTable.Rows[i][c], compareDataTable.Rows[i][c]))
                        return false;
                }
            }
            return true;
        }

        private string GetLastModifiedFile(string path)
        {
            var files =  new DirectoryInfo(path).GetFiles()
                        .OrderByDescending(f => f.LastWriteTime)
                        .Select(f => f.FullName)
                        .ToList();
            return files.FirstOrDefault();
        }

        private bool CompareFile(string originalFile, string compareFile, bool rearrange = false)
        {
            ExcelReader reader = new ExcelReader();
            string originalFileActiveSheetName = "";
            string compareFileActiveSheetName = "";
            originalFile = CleanString(originalFile);
            compareFile = CleanString(compareFile);
            var dsOriginalFile = reader.ExcelToDataSet(originalFile, "", -1, out originalFileActiveSheetName);
            var dsCompareFile = reader.ExcelToDataSet(compareFile, "", -1, out compareFileActiveSheetName);

            var originalDt = dsOriginalFile.Tables[0];
            var compareDt = dsCompareFile.Tables[0];
            if (rearrange)
            {
                originalDt = ReArrangeTable(originalDt);
                compareDt = ReArrangeTable(compareDt);
            }

            return CompareDataTable(originalDt, compareDt);
        }

        private string CleanString(string originalString)
        {
            var regex = new Regex(@"[\p{Cc}\p{Cf}\p{Mn}\p{Me}\p{Zl}\p{Zp}]");
            return regex.Replace(originalString, "");
        }

        private DataTable ReArrangeTable(DataTable locationTable)
        {
            // Create DataView
            DataView view = new DataView(locationTable);

            // Sort by column in asc order
            var sortColumnString = string.Empty;
            foreach (var columnName in locationTable.Columns)
            {
                sortColumnString += columnName + " ASC,";
            }
            sortColumnString = sortColumnString.TrimEnd(',');

            view.Sort = sortColumnString;
            return view.ToTable();
        }
    }
}
