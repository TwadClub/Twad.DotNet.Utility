using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace Twad.DotNet.Utility.Excel
{
    public partial class NPOIHelper
    {
        /// <summary>
        /// 战略协议单元格样式
        /// </summary>
        /// <param name="dicTitle"></param>
        /// <param name="excelWorkbook">excelWorkbook</param>
        /// <returns></returns>
        public static List<ExeclCellStyle> GetTacticCgAgreementExeclCellStyleList(Dictionary<string, string> dicTitle, IWorkbook excelWorkbook)
        {
            List<ExeclCellStyle> execlCellStyleList = new List<ExeclCellStyle>();
            foreach (var item in dicTitle)
            {
                string columnsName = item.Key.ToString();
                ExeclCellStyle execlCellStyleEntity = new ExeclCellStyle();
                execlCellStyleEntity.IsLock = false;
                execlCellStyleEntity.ColumnsName = columnsName;
                execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                execlCellStyleEntity.IsHidden = false;
                execlCellStyleEntity.TitleStyle = NoEditTitleStyle(excelWorkbook);
                execlCellStyleEntity.CellStyle = NoEditStyle(excelWorkbook);
                switch (columnsName)
                {
                    case "ProductCode":
                        execlCellStyleEntity.Width = 20; break;
                    case "ProdcutName":
                        execlCellStyleEntity.Width = 30; break;
                    case "ProductTypeName":
                        execlCellStyleEntity.Width = 10; break;
                    case "ProductAttribute":
                        execlCellStyleEntity.Width = 80; break;
                    case "ProductUnit":
                        execlCellStyleEntity.Width = 10;
                        execlCellStyleEntity.IsHidden = false; break;
                    case "TacticCgAgreementPrice":
                        execlCellStyleEntity.Width = 10;
                        execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.White.Index;
                        execlCellStyleEntity.TitleStyle = EditTitleStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle = EditStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle.Alignment = HorizontalAlignment.Right;//金额右对齐
                        break;
                    case "TacticCgAgreementProductRemark":
                        execlCellStyleEntity.Width = 80;
                        execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.White.Index;
                        execlCellStyleEntity.TitleStyle = EditTitleStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle = EditStyle(excelWorkbook);
                        break;
                    case "TacticCgAgreementGUID":
                        execlCellStyleEntity.Width = 5;
                        execlCellStyleEntity.IsHidden = true; break;
                    case "ProductGUID":
                        execlCellStyleEntity.Width = 5;
                        execlCellStyleEntity.IsHidden = true; break;
                    default:
                        break;
                }
                execlCellStyleList.Add(execlCellStyleEntity);
            }
            return execlCellStyleList;
        }


        /// <summary>
        /// 合同单元格样式
        /// </summary>
        /// <param name="dicTitle"></param>
        /// <param name="excelWorkbook">excelWorkbook</param>
        /// <returns></returns>
        public static List<ExeclCellStyle> GetContractExeclCellStyleList(Dictionary<string, string> dicTitle, IWorkbook excelWorkbook)
        {
            List<ExeclCellStyle> execlCellStyleList = new List<ExeclCellStyle>();
            foreach (var item in dicTitle)
            {
                string columnsName = item.Key.ToString();
                ExeclCellStyle execlCellStyleEntity = new ExeclCellStyle();
                execlCellStyleEntity.IsLock = false;
                execlCellStyleEntity.ColumnsName = columnsName;
                execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                execlCellStyleEntity.IsHidden = false;
                execlCellStyleEntity.TitleStyle = NoEditTitleStyle(excelWorkbook);
                execlCellStyleEntity.CellStyle = NoEditStyle(excelWorkbook);
                switch (columnsName)
                {
                    case "ProductCode":
                        execlCellStyleEntity.Width = 20; break;
                    case "ProdcutName":
                        execlCellStyleEntity.Width = 30; break;
                    case "ProductTypeName":
                        execlCellStyleEntity.Width = 10; break;
                    case "ProductAttribute":
                        execlCellStyleEntity.Width = 80; break;
                    case "ProductUnit":
                        execlCellStyleEntity.Width = 10; break;
                    case "TacticCgAgreementPrice":
                        execlCellStyleEntity.Width = 30;
                        execlCellStyleEntity.CellStyle.Alignment = HorizontalAlignment.Right;//金额右对齐
                        break;
                    case "ContractCount":
                        execlCellStyleEntity.Width = 10;
                        execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.White.Index;
                        execlCellStyleEntity.TitleStyle = EditTitleStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle = EditStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle.Alignment = HorizontalAlignment.Right;//金额右对齐
                        break;
                    case "ContractPrice":
                        execlCellStyleEntity.Width = 15;
                        execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.White.Index;
                        execlCellStyleEntity.TitleStyle = EditTitleStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle = EditStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle.Alignment = HorizontalAlignment.Right;//金额右对齐
                        break;
                    case "TotalPrice":
                        execlCellStyleEntity.Width = 15;
                        execlCellStyleEntity.CellStyle.Alignment = HorizontalAlignment.Right;//金额右对齐
                        execlCellStyleEntity.SetCellFormula = "I{0}*J{0}";
                        break;
                    case "ContractProductRemark":
                        execlCellStyleEntity.Width = 80;
                        execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.White.Index;
                        execlCellStyleEntity.TitleStyle = EditTitleStyle(excelWorkbook);
                        execlCellStyleEntity.CellStyle = EditStyle(excelWorkbook);
                        break;
                    case "ContractGUID":
                        execlCellStyleEntity.Width = 5;
                        execlCellStyleEntity.IsHidden = true; break;
                    case "ProductGUID":
                        execlCellStyleEntity.Width = 5;
                        execlCellStyleEntity.IsHidden = true; break;
                    default:
                        break;
                }
                execlCellStyleList.Add(execlCellStyleEntity);
            }
            return execlCellStyleList;
        }


        /// <summary>
        /// 非编辑列头部样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle NoEditTitleStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(GetTitleFont(excelWorkbook));
            //设置背景色
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }
        /// <summary>
        /// 编辑列头部样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle EditTitleStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(GetTitleFont(excelWorkbook));
            //设置背景色
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightOrange.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }
        /// <summary>
        /// 编辑单元格样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle NoEditStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Left;
            //style.SetFont(GetTitleFont(excelWorkbook));
            //设置背景色
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }
        /// <summary>
        /// 非编辑单元格样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle EditStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Left;
            //style.SetFont(GetTitleFont(excelWorkbook));
            return style;
        }
        /// <summary>
        /// 头部字体
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static IFont GetTitleFont(IWorkbook excelWorkbook)
        {
            HSSFFont font = (HSSFFont)excelWorkbook.CreateFont();
            font.FontHeightInPoints = 11;//字号  
            font.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;//颜色 
            font.Boldweight = 500;//HSSFFont.BOLDWEIGHT_BOLD;//加粗  

            return font;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string GetWaring()
        {
            string msg ="exec警告";
            string[] str = msg.Split(new char[] { '.' });
            return string.Join("\n", str);//数组转成字符串 
        }
        /// <summary>
        /// 添加验证区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="validationAreaList">待验证区域列表</param>
        public static void SheetAddValidationArea(ISheet sheet, List<ValidationArea> validationAreaList)
        {
            if (validationAreaList == null || validationAreaList.Count <= 0)
            {
                return;
            }
            foreach (ValidationArea item in validationAreaList)
            {
                if (item.ValidationType == ValidationType.DECIMAL)
                {
                    SetCellInputNumber(sheet, item.FirstRow, item.LastRow, item.FirstCol, item.LastCol);
                }
                else if (item.ValidationType == ValidationType.ANY)
                {
                    //其他的类型的处理业务逻辑，需要的时候再补充
                }
            }
        }


        /// <summary>
        /// 设置只是输入数字
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="lastRow"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        public static void SetCellInputNumber(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            var cellRegions = new CellRangeAddressList(firstRow, lastRow, firstCol, firstCol);
            DVConstraint constraint = DVConstraint.CreateNumericConstraint(ValidationType.DECIMAL, OperatorType.BETWEEN, "0", "999999999");
            HSSFDataValidation dataValidate = new HSSFDataValidation(cellRegions, constraint);
            dataValidate.CreateErrorBox("", "经过语言");

            sheet.AddValidationData(dataValidate);
        }
    }
}
