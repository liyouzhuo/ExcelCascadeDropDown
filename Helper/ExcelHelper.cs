using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
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

        }
    }
}
