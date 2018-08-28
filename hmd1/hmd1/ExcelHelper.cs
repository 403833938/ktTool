using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.Sql;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace sxst
{
    class ExcelHelper
    {
        /// <summary>
        /// 由DataTable导出Excel
        /// </summary>
        /// <param name="sourceTable">要导出数据的DataTable</param>
        /// <returns>Excel工作表</returns>
        public void ExportDataTableToExcel(DataTable sourceTable, string sheetName, string filepath)
        {

            FileStream file = new FileStream(filepath, FileMode.Create);
            HSSFWorkbook workbook = new HSSFWorkbook();
            // MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet(sheetName);
            IRow headerRow = sheet.CreateRow(0);
            // handling header.
            foreach (DataColumn column in sourceTable.Columns)

                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
            // handling value.
            int rowIndex = 1;
            foreach (DataRow row in sourceTable.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in sourceTable.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }
                rowIndex++;
            }
            workbook.Write(file);
            file.Close();
            sheet = null;
            headerRow = null;
            workbook = null;
        }

    }   

}
