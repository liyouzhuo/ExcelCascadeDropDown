using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace Helper
{
    public class ExcelHelper
    {
        public void ExportExcel()
        {
            // 创建excel工作簿
            IWorkbook workbook = new XSSFWorkbook();
            // 创建单元格样式
            ICellStyle cellStyle = workbook.CreateCellStyle();

            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            // 水平对齐
            cellStyle.Alignment = HorizontalAlignment.Center;
            // 垂直对齐
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            // 设置字体
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 18;
            font.FontName = "微软雅黑";
            cellStyle.SetFont(font);

            // 创建sheet
            ISheet sheet = workbook.CreateSheet("Sheet1");
            // 设置第一列的宽度
            sheet.SetColumnWidth(0, 15 * 256);
            // 创建一行
            IRow row_0 = sheet.CreateRow(0);
            // 创建两个单元格
            ICell cell_0 = row_0.CreateCell(0);
            ICell cell_1 = row_0.CreateCell(1);

            cell_0.SetCellValue("省份");
            cell_0.CellStyle = cellStyle;
            cell_1.SetCellValue("城市");
            cell_1.CellStyle = cellStyle;

            using (FileStream fileStream = File.OpenWrite($"{Environment.CurrentDirectory}\\demo.xlsx"))
            {
                workbook.Write(fileStream);
            }
        }

        /// <summary>
        /// 格式化数据，并建立名称管理
        /// </summary>
        /// <param name="sheet">表</param>
        /// <param name="model">数据源(数据库)</param>
        /// <param name="firstCellName">第一列的名称</param>
        /// <param name="rowNo">当前操作的行号</param>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetName">工作表名</param>
        public void FormatData(ISheet sheet, List<string> model, string firstCellName, ref int rowNo, IWorkbook workbook, string sheetName)
        {
            // 按行写入类型数据
            IRow row = sheet.CreateRow(rowNo);
            row.CreateCell(0).SetCellValue(firstCellName);
            int rowCell = 1;
            foreach (var item in model)
            {
                row.CreateCell(rowCell).SetCellValue(item);
                rowCell++;
            }
            // 建立名称管理
            rowNo++;
            IName range = workbook.CreateName();
            range.NameName = firstCellName;
            string colName = GetExcelColumnName(model.Count + 1);
            range.RefersToFormula = string.Format("{0}!$B${1}:${2}${1}", sheetName, rowNo, colName);
            range.Comment = rowNo.ToString("00");
        }

        /// <summary>
        /// 获取Excel列名
        /// </summary>
        /// <param name="columnNumber">列的序号，如：A、B、C、AA、BB</param>
        /// <returns></returns>
        static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }

        /// <summary>
        /// 建立级联关系
        /// </summary>
        /// <param name="sheet">表</param>
        /// <param name="source">数据源(EXCEL表)</param>
        /// <param name="minRow">起始行</param>
        /// <param name="maxRow">终止行</param>
        /// <param name="minCell">起始列</param>
        /// <param name="maxCell">终止列</param>
        public void ExcelLevelRelation(ISheet sheet, string source, int minRow, int maxRow, int minCell, int maxCell)
        {
            // 第一层绑定下拉的时候，可以一次性选择多个单元格进行绑定
            // 要是从第二层开始，就只能一对一的绑定，如果目标单元格要与哪一个一级单元格进行关联
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet as XSSFSheet);
            XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)dvHelper.CreateFormulaListConstraint(source);
            CellRangeAddressList cellRegions = new CellRangeAddressList(minRow, maxRow, minCell, maxCell);
            XSSFDataValidation validation = (XSSFDataValidation)dvHelper.CreateValidation(dvConstraint, cellRegions);
            validation.SuppressDropDownArrow = true;
            validation.CreateErrorBox("输入不合法", "请选择下拉列表中的值。");
            validation.ShowErrorBox = true;
            sheet.AddValidationData(validation);
        }
    }
}
