using System;
using System.Data;
using System.Web;
using System.IO;
using System.Text;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace DHCC.HR.Common
{
    /// <summary>
    /// Excel操作类
    /// chinaway
    /// 2016-01-11
    /// </summary>
    public sealed class ExcelOperate
    {
        /// <summary>
        /// DataTable导出到excel
        /// </summary>
        /// <param name="dtSource">数据源</param>
        /// <param name="dicColAliasNames">导出的列重命名，可选</param>
        /// <param name="sFileName">文件名(包含后缀名)，可选</param>
        /// <param name="sSheetName">工作薄名，可选</param>
        public static void ToExcel(DataTable dtSource, IDictionary<string, string> dicColAliasNames = null, string sFileName = "新导出工作表.xls", string sSheetName = "Sheet")
        {
            HttpContext curContext = HttpContext.Current;
            // 设置编码和附件格式      
            curContext.Response.ContentType = "application/vnd.ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                "attachment;filename=" + HttpUtility.UrlEncode(sFileName, Encoding.UTF8));

            if (string.IsNullOrWhiteSpace(sFileName))
            {
                sFileName = "新导出工作表.xls";
            }
            if (string.IsNullOrWhiteSpace(sSheetName))
            {
                sSheetName = "Sheet";
            }
            bool isCompatible = Common.GetIsCompatible(sFileName);

            IWorkbook workbook = Common.CreateWorkbook(isCompatible);
            ICellStyle headerCellStyle = Common.GetCellStyle(workbook, true);
            ICellStyle cellStyle = Common.GetCellStyle(workbook);
            ISheet sheet = workbook.CreateSheet(sSheetName);
            int rowIndex = 1;
            int colIndex = 1;
            int rowIndexMax = 1048575;
            int colIndexMan = 16383;
            if (isCompatible)
            {
                rowIndexMax = 65535;
                colIndexMan = 255;
            }
            #region 创建列头
            IRow headerRow = sheet.CreateRow(0);
            if (dicColAliasNames == null || dicColAliasNames.Count == 0)
            {
                foreach (DataColumn column in dtSource.Columns)
                {
                    if (colIndex < colIndexMan)
                    {
                        ICell headerCell = headerRow.CreateCell(column.Ordinal);
                        headerCell.SetCellValue(column.ColumnName);
                        headerCell.CellStyle = headerCellStyle;
                        sheet.AutoSizeColumn(headerCell.ColumnIndex);
                        colIndex++;
                    }
                }
            }
            else
            {
                int i = 0;
                foreach (var dic in dicColAliasNames)
                {                    
                    if (i < colIndexMan)
                    {
                        ICell headerCell = headerRow.CreateCell(i);
                        headerCell.SetCellValue(dic.Value);
                        headerCell.CellStyle = headerCellStyle;
                        sheet.AutoSizeColumn(headerCell.ColumnIndex);
                        i++;
                    }
                }
            }
            #endregion
            #region 填充内容
            foreach (DataRow row in dtSource.Rows)
            {
                if (rowIndex % rowIndexMax == 0)
                {
                    sheet = workbook.CreateSheet(sSheetName + ((int)rowIndex / rowIndexMax).ToString());
                }
                IRow dataRow = sheet.CreateRow(rowIndex);
                if (dicColAliasNames == null || dicColAliasNames.Count == 0)
                {
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        ICell cell = dataRow.CreateCell(column.Ordinal);
                        string rowValue = (row[column] ?? "").ToString();
                        switch (column.DataType.ToString())
                        {
                            case "System.DateTime"://日期类型                                  
                                cell.SetCellValue(rowValue == "" ? rowValue : DateTime.Parse(rowValue).ToShortDateString());
                                break;
                            default:
                                cell.SetCellValue(rowValue);
                                break;
                        }
                        cell.CellStyle = cellStyle;
                        Common.ReSizeColumnWidth(sheet, cell);
                    }
                }
                else
                {
                    int i = 0;
                    foreach (var dic in dicColAliasNames)
                    {                        
                        ICell cell = dataRow.CreateCell(i);
                        string rowValue = (row[dtSource.Columns[dic.Key].Ordinal] ?? "").ToString();
                        switch (dtSource.Columns[dic.Key].DataType.ToString())
                        {
                            case "System.DateTime"://日期类型                                  
                                cell.SetCellValue(rowValue == "" ? rowValue : DateTime.Parse(rowValue).ToShortDateString());
                                break;
                            default:
                                cell.SetCellValue(rowValue);
                                break;
                        }
                        cell.CellStyle = cellStyle;
                        Common.ReSizeColumnWidth(sheet, cell);
                        i++;
                    }
                }
                rowIndex++;
            }
            #endregion
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Dispose();
                sheet = null;
                workbook = null;

                curContext.Response.BinaryWrite(ms.GetBuffer());
                curContext.Response.End();
            }
        }

        /// <summary>
        /// Excel导入到DataTable
        /// </summary>
        /// <param name="excelFilePath">Excel文件路径，为物理路径</param>
        /// <param name="sSheetName">Excel工作表名称，可选</param>
        /// <param name="headerRowIndex">Excel表头行索引，可选</param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable(string excelFilePath, string sSheetName = "Sheet1", int headerRowIndex = 0)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(sSheetName))
            {
                sSheetName = "Sheet1";
            }
            using (FileStream stream = File.OpenRead(excelFilePath))
            {
                bool isCompatible = Common.GetIsCompatible(excelFilePath);
                IWorkbook workbook = Common.CreateWorkbook(isCompatible, stream);
                ISheet sheet = workbook.GetSheet(sSheetName);
                DataTable table = Common.GetDataTableFromSheet(sheet, headerRowIndex);

                stream.Close();
                workbook = null;
                sheet = null;
                ClearNullRow(table);
                return table;
            }
        }

        /// <summary>
        /// Excel导入到DataSet，如果有多个工作表，则导入多个DataTable
        /// </summary>
        /// <param name="excelFilePath">Excel文件路径，为物理路径</param>
        /// <param name="headerRowIndex">Excel表头行索引，可选</param>
        /// <returns>DataSet</returns>
        public static DataSet ToDataSet(string excelFilePath, int headerRowIndex = 0)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                return null;
            }
            using (FileStream stream = File.OpenRead(excelFilePath))
            {
                DataSet ds = new DataSet();
                bool isCompatible = Common.GetIsCompatible(excelFilePath);
                IWorkbook workbook = Common.CreateWorkbook(isCompatible, stream);
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    DataTable table = Common.GetDataTableFromSheet(sheet, headerRowIndex);
                    ClearNullRow(table);

                    ds.Tables.Add(table);
                }
                stream.Close();
                workbook = null;

                return ds;
            }
        }

        /// <summary>
        /// 清空DataTable中的空行
        /// </summary>
        /// <param name="dtSource"></param>
        private static void ClearNullRow(DataTable dtSource)
        {
            for (int i = dtSource.Rows.Count - 1; i > 0; i--)
            {
                bool isNull = true;
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    if (dtSource.Rows[i][j] != null)
                    {
                        if (dtSource.Rows[i][j].ToString() != "")
                        {
                            isNull = false;
                            break;
                        }
                    }
                }
                if (isNull)
                {
                    dtSource.Rows[i].Delete();
                }
            }
        }
    }
}
