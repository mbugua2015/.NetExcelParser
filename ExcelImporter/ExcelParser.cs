using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;

namespace ExcelImporter
{
    public class ExcelParser
    {

        public static XSSFWorkbook LoadExcelFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException("file path cannot be null or empty");

            if (!File.Exists(filePath))
            {
                throw new ArgumentException($"File {filePath} does not exist");
            }

            XSSFWorkbook workbook;
            try
            {
                using(var fileStream= new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fileStream);
                    return workbook;
                }
            }
            catch(Exception ex)
            {                
                throw ex;
            }
        }

        public static XSSFSheet[] GetWorksheets(XSSFWorkbook workBook)
        {
            if (workBook == null) throw new ArgumentNullException("workbook cannot be null");

            XSSFSheet[] workSheets = new XSSFSheet[workBook.Count];

            for(int i = 0; i < workBook.Count; i++)
            {
                workSheets[i] = (XSSFSheet)workBook.GetSheet(workBook.GetSheetAt(i).SheetName);
            }

            return workSheets;
        }

        public static object[][] GetWorksheetValues(XSSFSheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException("worksheet cannot be null");

            int columns = worksheet.GetRow(0).Cells.Count;
            List<IRow> rows = new List<IRow>();

            int i = 0;
            while (worksheet.GetRow(i) != null)
            {
                rows.Add(worksheet.GetRow(i));
            }

            object[][] sheetValues = new object[rows.Count][];

            for(int j = 0; j < rows.Count; j++)
            {
                for(int k = 0; k < rows[j].Cells.Count; k++)
                {
                    var cell = rows[i].GetCell(j);

                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                sheetValues[i][j] = cell.NumericCellValue;
                                break;
                            case CellType.String:
                                sheetValues[i][j] = cell.StringCellValue;
                                break;
                        }
                    }                   
                }
            }

            return sheetValues;
        }
       
    }
}
