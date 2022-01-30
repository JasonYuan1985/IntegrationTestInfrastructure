using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Data;

namespace IntegrationAutomation
{
    public class ExcelReader
    {
        private IWorkbook GetWorkbook(string fileName)
        {
            IWorkbook workbook = null;
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                if (fileName.ToLower().IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.ToLower().IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                fs.Close();
            }

            return workbook;
        }

        public DataSet ExcelToDataSet(string fileName, string sheetName, int headerRowIndex, out string activeSheetName)
        {
            ISheet sheet = null;
            IWorkbook workbook = GetWorkbook(fileName);

            activeSheetName = workbook.GetSheetName(workbook.ActiveSheetIndex);

            DataSet ds = new DataSet();

            int headRowIndex = headerRowIndex;

            if (string.IsNullOrEmpty(sheetName) == true)
            {

                ArrayList sheetNameList = GetSheetName(workbook);
                foreach (string sheetName1 in sheetNameList)
                {
                    sheet = workbook.GetSheet(sheetName1);
                    DataTable dt = ImportDataTable(sheet, headRowIndex, true);
                    dt.TableName = sheetName1;
                    ds.Tables.Add(dt);
                }
            }
            else
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                {
                    sheet = workbook.GetSheetAt(0);
                }
                DataTable dt = ImportDataTable(sheet, headRowIndex, true);
                dt.TableName = sheetName;
                ds.Tables.Add(dt);
            }
            return ds;
        }

        public ArrayList GetSheetName(IWorkbook workbook)
        {
            ArrayList arrayList = new ArrayList();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                arrayList.Add(workbook.GetSheetName(i));
            }
            return arrayList;
        }

        /// <summary>
        /// 将制定sheet中的数据导出到datatable中
        /// </summary>
        /// <param name="sheet">需要导出的sheet</param>
        /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
        /// <returns></returns>
        private DataTable ImportDataTable(ISheet sheet, int HeaderRowIndex = -1, bool needHeader = true)
        {
            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            if (HeaderRowIndex < 0 || !needHeader)
            {
                IRow detailRow = null;
                int firstCellNum, maxCellNum;
                firstCellNum = 0;
                maxCellNum = 0;
                for (int k = sheet.FirstRowNum; k < sheet.LastRowNum; k++)
                {
                    detailRow = sheet.GetRow(k) as IRow;
                    if (detailRow != null)
                    {
                        maxCellNum = maxCellNum > detailRow.LastCellNum ? maxCellNum : detailRow.LastCellNum;
                    }
                }

                cellCount = maxCellNum;
                for (int i = firstCellNum; i <= cellCount; i++)
                {
                    DataColumn column = new DataColumn(Convert.ToString(i));
                    table.Columns.Add(column);
                }
            }
            else
            {
                headerRow = sheet.GetRow(HeaderRowIndex) as IRow;
                cellCount = headerRow.LastCellNum;

                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    if (headerRow.GetCell(i) == null)
                    {
                        if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                        {
                            DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn(Convert.ToString(i));
                            table.Columns.Add(column);
                        }

                    }
                    else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                    {
                        DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                        table.Columns.Add(column);
                    }
                    else
                    {
                        DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                        table.Columns.Add(column);
                    }
                }
            }

            var rowCount = sheet.LastRowNum;
            for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row;
                if (sheet.GetRow(i) == null)
                {
                    row = sheet.CreateRow(i) as IRow;
                }
                else
                {
                    row = sheet.GetRow(i) as IRow;
                }

                DataRow dataRow = table.NewRow();

                if (row.Cells.Count == 0)
                {
                    table.Rows.Add(dataRow);
                    continue;
                }
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        dataRow[j] = GetRowValue(row, j);
                    }
                }
                table.Rows.Add(dataRow);
            }
            return table;
        }

        private object? GetRowValue(IRow row, int j)
        {
            object? rowContent;
            switch (row.GetCell(j).CellType)
            {
                case CellType.String:
                    string str = row.GetCell(j).StringCellValue;
                    if (str != null && str.Length > 0)
                    {
                        rowContent = str.ToString();
                    }
                    else
                    {
                        rowContent = null;
                    }
                    break;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                    {
                        rowContent = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                    }
                    else
                    {
                        rowContent = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                    }
                    break;
                case CellType.Boolean:
                    rowContent = Convert.ToString(row.GetCell(j).BooleanCellValue);
                    break;
                case CellType.Error:
                    rowContent = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                    break;
                case CellType.Formula:
                    switch (row.GetCell(j).CachedFormulaResultType)
                    {
                        case CellType.String:
                            string strFORMULA = row.GetCell(j).StringCellValue;
                            if (strFORMULA != null && strFORMULA.Length > 0)
                            {
                                rowContent = strFORMULA.ToString();
                            }
                            else
                            {
                                rowContent = null;
                            }
                            break;
                        case CellType.Numeric:
                            rowContent = Convert.ToString(row.GetCell(j).NumericCellValue);
                            break;
                        case CellType.Boolean:
                            rowContent = Convert.ToString(row.GetCell(j).BooleanCellValue);
                            break;
                        case CellType.Error:
                            rowContent = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                            break;
                        default:
                            rowContent = "";
                            break;
                    }
                    break;
                default:
                    rowContent = "";
                    break;
            }

            return rowContent;
        }
    }
}
