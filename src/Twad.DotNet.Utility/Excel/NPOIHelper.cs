using Microsoft.AspNetCore.Http;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace Twad.DotNet.Utility.Excel
{
    public partial class NPOIHelper
    {
        /// <summary>
        /// 需要初始化设置的表头的列名（必须设置）
        /// </summary>
        public static SortedList ListColumnsName;

        /// <summary>
        /// 最后一行总计列
        /// </summary>
        public static SortedList TotalColummsName;



        #region 导出

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        public static void ExportExcel(IList list, string filePath)
        {
            DataTable dtSource = DataHelper.ToDataTable(list);
            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));
            IWorkbook excelWorkbook = CreateExcelFile();
            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, filePath);
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        public static void ExportExcel(DataTable dtSource, string filePath)
        {
            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));

            IWorkbook excelWorkbook = CreateExcelFile();
            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, filePath);
        }


        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        public static void ExportExcel(IList list, Stream excelStream)
        {
            DataTable dtSource = DataHelper.ToDataTable(list);

            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));

            IWorkbook excelWorkbook = CreateExcelFile();
            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, excelStream);
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        public static void ExportExcel(DataTable dtSource, Stream excelStream)
        {
            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));

            IWorkbook excelWorkbook = CreateExcelFile();
            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, excelStream);
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        public static void ExportExcel(HttpResponse response, IList list, Stream excelStream, string fileName)
        {
            DataTable dtSource = DataHelper.ToDataTable(list);

            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));

            IWorkbook excelWorkbook = CreateExcelFile();

            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, excelStream);


            //response.AddHeader("content-disposition", string.Format("attachment;filename={0}({1}).xls", fileName, DateTime.Now.ToString("yyyyMMdd")));
            //response.Charset = "utf-8";
            //response.ContentEncoding = System.Text.Encoding.GetEncoding("gb2312");
            response.ContentType = "application/vnd.xls";
            response.Body = excelStream;
     

        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="response"></param>
        /// <param name="list"></param>
        /// <param name="fileName"></param>
        public static void ExportExcel(HttpResponse response, IList list, string fileName)
        {
            MemoryStream excelStream = new MemoryStream();

            DataTable dtSource = DataHelper.ToDataTable(list);

            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));

            IWorkbook excelWorkbook = CreateExcelFile();

            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, excelStream);
            response.ContentType = "application/vnd.xls";
            response.Body=excelStream;

        


        }

     
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="response"></param>
        /// <param name="list"></param>
        /// <param name="fileName"></param>
        public static void ExportExcel(HttpResponse response, DataTable dtSource, string fileName)
        {
            MemoryStream excelStream = new MemoryStream();

            if (ListColumnsName == null || ListColumnsName.Count == 0)
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));

            IWorkbook excelWorkbook = CreateExcelFile();

            InsertRow(dtSource, excelWorkbook);
            SaveExcelFile(excelWorkbook, excelStream);
            response.ContentType = "application/vnd.xls";
            response.Body=excelStream;

        }
        #endregion

        #region 导入

        /// <summary>    
        /// 由Excel导入到DataTable    
        /// </summary>    
        /// <param name="excelFileStream">Excel文件流</param>    
        /// <param name="sheetName">Excel工作表名称</param>    
        /// <param name="headerRowIndex">Excel表头行索引</param>    
        /// <returns>DataTable</returns>    
        public static DataTable ImportExcel(Stream excelFileStream, int sheetIndex, int headerRowIndex)
        {
            try
            {
                IWorkbook workbook;
                try
                {
                    workbook = new XSSFWorkbook(excelFileStream);//2007版本
                }
                catch
                {
                    workbook = new HSSFWorkbook(excelFileStream);//2003版本
                }

                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                DataTable table = new DataTable();
                IRow headerRow = sheet.GetRow(headerRowIndex);
                int cellCount = headerRow.LastCellNum;
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    table.Columns.Add(column);
                }
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null)
                        continue;

                    DataRow dataRow = table.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        dataRow[j] = row.GetCell(j) == null ? "" : row.GetCell(j).ToString().Trim();
                    }

                    //判断正行数据是否有效
                    bool key = false;
                    for (int ii = 0; ii < cellCount; ii++)
                    {
                        if (!string.IsNullOrEmpty(dataRow[ii].ToString()))
                        {
                            key = true;
                            break;
                        }
                    }
                    if (key)
                    {
                        table.Rows.Add(dataRow);
                    }
                }
                workbook = null;
                sheet = null;
                return table;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                excelFileStream.Close();
            }
        }

        #endregion

        #region private

        /// <summary>
        /// CellBorder设置（表头）
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        private static ICellStyle BorderCellStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.FillBackgroundColor = (short)999;
            return style;
        }
        /// <summary>
        /// 保存Excel文件
        /// </summary>
        /// <param name="excelWorkBook"></param>
        /// <param name="filePath"></param>
        private static void SaveExcelFile(IWorkbook excelWorkBook, string filePath)
        {
            FileStream file = null;
            try
            {
                file = new FileStream(filePath, FileMode.Create);
                excelWorkBook.Write(file);
            }
            finally
            {
                if (file != null)
                {
                    file.Close();
                }
            }
        }
        /// <summary>
        /// 保存Excel文件
        /// </summary>
        /// <param name="excelWorkBook"></param>
        /// <param name="filePath"></param>
        private static void SaveExcelFile(IWorkbook excelWorkBook, Stream excelStream)
        {
            try
            {
                excelWorkBook.Write(excelStream);
            }
            finally
            {

            }
        }
        /// <summary>
        /// 创建Excel文件
        /// </summary>
        /// <param name="filePath"></param>
        private static IWorkbook CreateExcelFile()
        {
            IWorkbook hssfworkbook = new HSSFWorkbook();
            return hssfworkbook;
        }
        /// <summary>
        /// 创建excel表头
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="excelSheet"></param>
        private static void CreateHeader(ISheet excelSheet, ICellStyle style)
        {
            int cellIndex = 0;
            IRow newRow = excelSheet.CreateRow(0);
            //循环导出列
            foreach (DictionaryEntry de in ListColumnsName)
            {
                ICell newCell = newRow.CreateCell(cellIndex);
                newCell.SetCellValue(de.Value.ToString());
                //按表头文字的宽度
                excelSheet.SetColumnWidth(cellIndex, de.Value.ToString().Length * 650);
                newCell.CellStyle = style;
                cellIndex++;
            }
        }
        /// <summary>
        /// 创建Total列
        /// </summary>
        /// <param name="excelSheet"></param>
        private static void CreateTotalFooter(ISheet excelSheet, int lastRow)
        {
            //无Total列
            if (TotalColummsName == null || TotalColummsName.Count == 0)
                return;

            string firstTotalColumnName = (TotalColummsName.GetKeyList()[0]).ToString();

            int cellIndex = 0;
            foreach (DictionaryEntry de in ListColumnsName)
            {
                if (de.Key.ToString() == firstTotalColumnName)
                {
                    break;
                }
                cellIndex++;
            }

            if (cellIndex == -1)
                return;

            IRow newRow = excelSheet.CreateRow(lastRow);

            //循环导出Total列
            foreach (DictionaryEntry footColumn in TotalColummsName)
            {
                ICell newCell = newRow.CreateCell(cellIndex);
                newCell.SetCellValue(footColumn.Value.ToString());
                cellIndex++;
            }

        }

        /// <summary>
        /// 插入数据行
        /// </summary>
        private static void InsertRow(DataTable dtSource, IWorkbook excelWorkbook)
        {
            int rowCount = 0;
            int sheetCount = 1;
            ISheet newsheet = null;

            //循环数据源导出数据集
            newsheet = excelWorkbook.CreateSheet("Sheet" + sheetCount);

            ICellStyle style = BorderCellStyle(excelWorkbook);

            //设置表头
            CreateHeader(newsheet, style);

            foreach (DataRow dr in dtSource.Rows)
            {
                rowCount++;
                //超出65530条数据 创建新的工作簿
                if (rowCount == 65530)
                {
                    rowCount = 1;
                    sheetCount++;
                    newsheet = excelWorkbook.CreateSheet("Sheet" + sheetCount);
                    CreateHeader(newsheet, style);
                }

                IRow newRow = newsheet.CreateRow(rowCount);
                InsertCell(dtSource, dr, newRow, newsheet, excelWorkbook);
            }

            //设置最后一列：Total列(最后一个sheet有总计)
            CreateTotalFooter(newsheet, dtSource.Rows.Count + 1);

        }
        /// <summary>
        /// 导出数据行
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="drSource"></param>
        /// <param name="currentExcelRow"></param>
        /// <param name="excelSheet"></param>
        /// <param name="excelWorkBook"></param>
        private static void InsertCell(DataTable dtSource, DataRow drSource, IRow currentExcelRow, ISheet excelSheet, IWorkbook excelWorkBook)
        {
            for (int cellIndex = 0; cellIndex < ListColumnsName.Count; cellIndex++)
            {
                //列名称
                string columnsName = ListColumnsName.GetKey(cellIndex).ToString();
                ICell newCell = null;
                System.Type rowType = drSource[columnsName].GetType();
                string drValue = drSource[columnsName].ToString().Trim();
                switch (rowType.ToString())
                {
                    case "System.String"://字符串类型
                        drValue = drValue.Replace("&", "&");
                        drValue = drValue.Replace(">", ">");
                        drValue = drValue.Replace("<", "<");
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(drValue);
                        break;
                    case "System.DateTime"://日期类型
                        DateTime dateV;
                        DateTime.TryParse(drValue, out dateV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(dateV);

                        //格式化显示
                        ICellStyle cellStyle = excelWorkBook.CreateCellStyle();
                        IDataFormat format = excelWorkBook.CreateDataFormat();
                        cellStyle.DataFormat = format.GetFormat("yyyy-mm-dd hh:mm:ss");
                        newCell.CellStyle = cellStyle;

                        break;
                    case "System.Boolean"://布尔型
                        bool boolV = false;
                        bool.TryParse(drValue, out boolV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(boolV);
                        break;
                    case "System.Int16"://整型
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Byte":
                        int intV = 0;
                        int.TryParse(drValue, out intV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(intV.ToString());
                        break;
                    case "System.Decimal"://浮点型
                    case "System.Double":
                        double doubV = 0;
                        double.TryParse(drValue, out doubV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(doubV);
                        break;
                    case "System.DBNull"://空值处理
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue("");
                        break;
                    default:
                        throw (new Exception(rowType.ToString() + "：类型数据无法处理!"));
                }
            }
        }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="response">返回流</param>
        /// <param name="excelWorkbook">excelWorkbook</param>
        /// <param name="list">数据集合</param>
        /// <param name="execlCellStyleList">单元格样式集合</param>
        /// <param name="validationAreaList">格式验证区域</param>
        /// <param name="fileName">文件名称</param>
        public static void ExportExcel(HttpResponse response, IWorkbook excelWorkbook, IList list, List<ExeclCellStyle> execlCellStyleList, List<ValidationArea> validationAreaList, string fileName)
        {
            if (ListColumnsName == null || ListColumnsName.Count == 0)
            {
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));
            }
            MemoryStream excelStream = ExportExcel(excelWorkbook, list, execlCellStyleList, validationAreaList);
     

            response.ContentType = "application/vnd.xls";
            response.Body= excelStream;
     
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelWorkbook">excelWorkbook</param>
        /// <param name="list">数据集合</param>
        /// <param name="execlCellStyleList">单元格样式集合</param>
        /// <param name="validationAreaList">格式验证区域</param>
        /// <returns></returns>
        public static MemoryStream ExportExcel(IWorkbook excelWorkbook, IList list, List<ExeclCellStyle> execlCellStyleList, List<ValidationArea> validationAreaList)
        {
            MemoryStream excelStream = new MemoryStream();
            DataTable dtSource = DataHelper.ToDataTable(list);
            InsertRow(dtSource, excelWorkbook, execlCellStyleList, validationAreaList);
            SaveExcelFile(excelWorkbook, excelStream);
            return excelStream;
        }
        /// <summary>
        /// 插入数据行
        /// </summary>
        /// <param name="dtSource">数据表</param>
        /// <param name="excelWorkbook">工作簿</param>
        /// <param name="execlCellStyleList">execl单元格样式</param>
        /// <param name="validationAreaList">格式验证区域</param>
        private static void InsertRow(DataTable dtSource, IWorkbook excelWorkbook, List<ExeclCellStyle> execlCellStyleList, List<ValidationArea> validationAreaList)
        {
            int rowCount = 1;
            int sheetCount = 1;
            ISheet newsheet = null;

            //循环数据源导出数据集
            newsheet = excelWorkbook.CreateSheet("Sheet" + sheetCount);

            ICellStyle style = BorderCellStyle(excelWorkbook);
            //合并单元格
            SetCellRangeAddress(newsheet, 0, 0, 1, 6);
            IRow woringRow = newsheet.CreateRow(0);
            woringRow.Height = 30 * 20;
            ICell newCell = woringRow.CreateCell(1);
            newCell.CellStyle = style;
            newCell.CellStyle.Alignment = HorizontalAlignment.Left;
            newCell.SetCellValue(GetWaring());
            newCell.CellStyle.WrapText = true;

            //设置表头
            CreateHeader(newsheet, style, execlCellStyleList);

            foreach (DataRow dr in dtSource.Rows)
            {
                rowCount++;
                IRow newRow = newsheet.CreateRow(rowCount);

                InsertCell(dtSource, dr, newRow, newsheet, excelWorkbook, execlCellStyleList, rowCount);
            }
            //当前工作薄添加数字限制区域
            SheetAddValidationArea(newsheet, validationAreaList);
        }
    
    private static void CreateHeader(ISheet excelSheet, ICellStyle style, List<ExeclCellStyle> execlCellStyleList)
    {
        int cellIndex = 0;
        IRow newRow = excelSheet.CreateRow(1);
        //循环导出列
        foreach (var de in ListColumnsName)
        {
            ICell newCell = newRow.CreateCell(cellIndex);
            newCell.SetCellValue(de.ToString());
            //按表头文字的宽度
            excelSheet.SetColumnWidth(cellIndex, de.ToString().Length * 650);
            ExeclCellStyle execlCellStyle = execlCellStyleList.FirstOrDefault(it => it.ColumnsName == de.ToString());
            if (execlCellStyle != null)
            {
                newCell.CellStyle = execlCellStyle.TitleStyle;//用户自定义表头样式
            }
            else
            {
                newCell.CellStyle = style;//赋默认值
            }
            cellIndex++;
        }
    }




    /// <summary>
    /// 导出数据行
    /// </summary>
    /// <param name="dtSource">数据表</param>
    /// <param name="drSource">来源数据行</param>
    /// <param name="currentExcelRow">当前数据行</param>
    /// <param name="excelSheet">sheet集合</param>
    /// <param name="excelWorkBook">数据表</param>
    /// <param name="execlCellStyleList">表头样式</param>
    /// <param name="rowCount">execl当前行</param>
    private static void InsertCell(DataTable dtSource, DataRow drSource, IRow currentExcelRow, ISheet excelSheet, IWorkbook excelWorkBook, List<ExeclCellStyle> execlCellStyleList, int rowCount)
        {
            int cellIndex = 0;
            foreach (var item in ListColumnsName)
            {
                //列名称
                string columnsName = item.ToString();
                //根据列名称设置列样式
                ExeclCellStyle execlCellStyle = execlCellStyleList?.FirstOrDefault(it => it.ColumnsName == columnsName) ?? null;
                ICellStyle columnStyle = excelSheet.GetColumnStyle(cellIndex);
                //设置列样式
                //excelSheet.SetDefaultColumnStyle(cellIndex, SetExeclColumnStyle(columnStyle, excelWorkBook, execlCellStyle));
                //设置列宽
                if ((execlCellStyle?.Width ?? 0) > 0)
                {
                    excelSheet.SetColumnWidth(cellIndex, execlCellStyle.Width);
                }
                if (execlCellStyle?.IsHidden ?? false)
                {
                    excelSheet.SetColumnHidden(cellIndex, true);
                }
                ICell newCell = null;
                System.Type rowType = drSource[columnsName].GetType();
                string drValue = drSource[columnsName].ToString().Trim();
                //错误消息列
                string errColumn = columnsName + "ErrMsg";
                string errMsg = string.Empty;
                if (dtSource.Columns.Contains(errColumn))
                {
                    errMsg = drSource[errColumn].ToString().Trim();
                }
                switch (rowType.ToString())
                {
                    case "System.String"://字符串类型
                        drValue = drValue.Replace("&", "&");
                        drValue = drValue.Replace(">", ">");
                        drValue = drValue.Replace("<", "<");
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(drValue);
                        break;
                    case "System.Guid"://GUID类型                        
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(drValue);
                        break;
                    case "System.DateTime"://日期类型
                        DateTime dateV;
                        DateTime.TryParse(drValue, out dateV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(dateV);
                        //格式化显示
                        ICellStyle cellStyle = excelWorkBook.CreateCellStyle();
                        IDataFormat format = excelWorkBook.CreateDataFormat();
                        cellStyle.DataFormat = format.GetFormat("yyyy-mm-dd hh:mm:ss");
                        newCell.CellStyle = cellStyle;
                        break;
                    case "System.Boolean"://布尔型
                        bool boolV = false;
                        bool.TryParse(drValue, out boolV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(boolV);
                        break;
                    case "System.Int16"://整型
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Byte":
                        int intV = 0;
                        int.TryParse(drValue, out intV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(intV.ToString());
                        break;
                    case "System.Decimal"://浮点型
                    case "System.Double":
                        double doubV = 0;
                        double.TryParse(drValue, out doubV);
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue(doubV);
                        break;
                    case "System.DBNull"://空值处理
                        newCell = currentExcelRow.CreateCell(cellIndex);
                        newCell.SetCellValue("");
                        break;
                    default:
                        throw (new Exception(rowType.ToString() + "：类型数据无法处理!"));
                }
                //设置单元格样式
                SetExeclCellStyle(newCell, execlCellStyle);
                //设置批注
                if (!string.IsNullOrWhiteSpace(errMsg))
                {
                    SetComment(errMsg, currentExcelRow, newCell, excelWorkBook.GetCreationHelper(), (HSSFPatriarch)excelSheet.CreateDrawingPatriarch());
                }
                //设置计算表达式
                if (!string.IsNullOrEmpty(execlCellStyle.SetCellFormula))
                {
                    int rowIndex = rowCount + 1;
                    newCell.SetCellFormula(string.Format(execlCellStyle.SetCellFormula, rowIndex));
                    if (excelSheet.ForceFormulaRecalculation == false)
                    {
                        excelSheet.ForceFormulaRecalculation = true;//没有此句，则不会刷新出计算结果
                    }
                }
                cellIndex++;
            }
            bool shouldLock = execlCellStyleList?.FirstOrDefault(it => it.IsLock == true)?.IsLock ?? false;
            if (shouldLock == true)
            {
                excelSheet.ProtectSheet("password");//设置密码保护
            }

        }
        /// <summary>
        /// 列锁定
        /// </summary>
        /// <param name="columnStyle"></param>
        /// <param name="excelWorkBook"></param>
        /// <param name="execlCellStyle"></param>
        private static ICellStyle SetExeclColumnStyle(ICellStyle columnStyle, IWorkbook excelWorkBook, ExeclCellStyle execlCellStyle)
        {
            if (columnStyle == null)
            {
                columnStyle = excelWorkBook.CreateCellStyle();
            }
            if (execlCellStyle != null)
            {
                columnStyle.IsLocked = execlCellStyle.IsLock;
                columnStyle.IsHidden = execlCellStyle.IsHidden;
            }
            return columnStyle;
        }
        /// <summary>
        /// 单元格格式
        /// </summary>
        /// <param name="execlCell"></param>
        /// <param name="execlCellStyle"></param>
        /// <returns></returns>
        private static void SetExeclCellStyle(ICell execlCell, ExeclCellStyle execlCellStyle)
        {
            execlCell.CellStyle = execlCellStyle.CellStyle;
        }
        /// <summary>
        /// 设置批注
        /// </summary>
        /// <param name="errMsg"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <param name="facktory"></param>
        /// <param name="patr"></param>
        public static void SetComment(string errMsg, IRow row, ICell cell, ICreationHelper facktory, HSSFPatriarch patr)
        {


            var anchor = facktory.CreateClientAnchor();
            anchor.Col1 = cell.ColumnIndex;
            anchor.Col2 = cell.ColumnIndex + 3;
            anchor.Row1 = row.RowNum;
            anchor.Row2 = row.RowNum + 5;
            var comment = patr.CreateCellComment(anchor);
            comment.String = new HSSFRichTextString(errMsg);
            comment.Author = ("mysoft");
            cell.CellComment = (comment);

        }
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="sheet">要合并单元格所在的sheet</param>
        /// <param name="rowstart">开始行的索引</param>
        /// <param name="rowend">结束行的索引</param>
        /// <param name="colstart">开始列的索引</param>
        /// <param name="colend">结束列的索引</param>
        public static void SetCellRangeAddress(ISheet sheet, int rowstart, int rowend, int colstart, int colend)
        {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowstart, rowend, colstart, colend);
            sheet.AddMergedRegion(cellRangeAddress);
        }

        #endregion
    }
    //排序实现接口 不进行排序 根据添加顺序导出
    public class NoSort : System.Collections.IComparer
    {
        public int Compare(object x, object y)
        {
            return -1;
        }
    }
}

