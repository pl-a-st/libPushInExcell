using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace libPushInExcell
{
    using NPOI;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.POIFS.Crypt;
    using NPOI.SS.Formula.Functions;
    using NPOI.XSSF.UserModel;
    using System.IO;
    using System.Threading;
    using NPOI.POIFS.Crypt.Dsig;

    public enum ToExcellResult
    {
        Unknown, Success, Failure, Exception, ParamError, AcsessError, Canceled
    }
    public enum ValueTypeExcell
    {
        String, Double, DateTime
    }
    public enum AddressFormat
    {
        A1, R1C1
    }
    /// <summary>
    ///   представляет упрощенный инструмент записи в файлы xlsx, xlsm через пакет NPOI
    /// </summary>
    public static class ToExcell
    {
        private static IWorkbook Workbook { get; set; }
        /// <summary>
        /// Полное имя файла для работы. Будет выбираться этот параметр, если в DataForExcell не указан адрес.
        /// </summary>
        public static string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                SetWorkbook(value);
                fileName = value;
            }
        }
        private static string fileName = string.Empty;
        /// <summary>
        /// Содержит последние сообщение об ошибках при неудачной записи в файл.
        /// </summary>
        public static string ExeptionMassage { get; private set; } = string.Empty;
        /// <summary>
        /// Формирует виртуальную книгу для 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static ToExcellResult SetWorkbook(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return ToExcellResult.Failure;
            }
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    Workbook = new XSSFWorkbook(fileStream);
                }
                return ToExcellResult.Success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"new FileStream({fileName}, FileMode.Open, FileAccess.ReadWrite)) вернул: " + ex.ToString());
                ExeptionMassage = ex.Message;
                return ToExcellResult.Failure;
            }
        }
        private static ToExcellResult SetWorkbook()
        {
            if (string.IsNullOrEmpty(FileName))
            {
                return ToExcellResult.Failure;
            }
            try
            {
                using (FileStream fileStream = new FileStream(FileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    Workbook = new XSSFWorkbook(fileStream);
                }
                return ToExcellResult.Success;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"new FileStream({FileName}, FileMode.Open, FileAccess.ReadWrite)) вернул: " + ex.ToString());
                ExeptionMassage = ex.Message;
                return ToExcellResult.Failure;
            }
        }
        private static ToExcellResult SaveWorkbook(string fileName)
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
                {
                    Workbook.Write(fileStream, false);
                }
                return ToExcellResult.Success;
            }
            catch (Exception ex)
            {
                ExeptionMassage = ex.Message;
                Console.WriteLine($"new FileStream({fileName}, FileMode.Create)) вернул: " + ex.ToString());
                return ToExcellResult.Failure;
            }
        }
        private static ToExcellResult SaveWorkbook()
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
                {
                    Workbook.Write(fileStream, false);
                }
                return ToExcellResult.Success;
            }
            catch (Exception ex)
            {
                ExeptionMassage = ex.Message;
                Console.WriteLine($"new FileStream({fileName}, FileMode.Create)) вернул: " + ex.ToString());
                return ToExcellResult.Failure;
            }
        }
        /// <summary>
        /// Записывает данные в эксель
        /// </summary>
        /// <param name="data"> данные для записи</param>
        /// <returns> ToExcellResult</returns>
        public static ToExcellResult Push(DataForExcell data)
        {
            try
            {
                Monitor.Enter(Workbook);
                if (!string.IsNullOrEmpty(data.FileName))
                {
                    SetWorkbook(data.FileName);
                }
                pushDataInWorkbook(data);
                SaveWorkbook(data);
                SetWorkbook();
            }
            catch (Exception exp)
            {
                ExeptionMassage = exp.Message;
                Console.WriteLine("Push(DataForExcell data) вернул " + exp.Message);
                return ToExcellResult.Exception;
            }
            finally
            {
                Monitor.Exit(Workbook);
            }
            return ToExcellResult.Success;
        }
        /// <summary>
        /// Записывает данные в эксель
        /// </summary>
        /// <param name="dataList"> Лист данных для записи</param>
        /// <returns> ToExcellResult</returns>
        public static ToExcellResult Push(List<DataForExcell> dataList)
        {
            try
            {
                Monitor.Enter(Workbook);
                foreach (DataForExcell data in dataList)
                {
                    if (!string.IsNullOrEmpty(data.FileName))
                    {
                        SetWorkbook(data.FileName);
                    }
                    pushDataInWorkbook(data);
                    SaveWorkbook(data);
                    SetWorkbook();
                }
            }
            catch (Exception exp)
            {
                ExeptionMassage = exp.Message;
                Console.WriteLine("Push(DataForExcell data) вернул " + exp.Message);
                return ToExcellResult.Exception;
            }
            finally
            {
                Monitor.Exit(Workbook);
            }
            return ToExcellResult.Success;
        }
        private static void SaveWorkbook(DataForExcell data)
        {
            if (!string.IsNullOrEmpty(data.FileName))
            {
                SaveWorkbook(data.FileName);
            }
            else
            {
                SaveWorkbook();
            }
        }

        private static void pushDataInWorkbook(DataForExcell data)
        {

            ISheet sheet = Workbook.GetSheetAt(data.SheetIndex);
            IRow row = sheet.GetRow(data.RowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(data.RowIndex);
            }
            ICell cell = row.GetCell(data.ColumnIndex);
            if (cell == null)
            {
                cell = row.CreateCell(data.ColumnIndex);
            }
            if (data.ValueType == ValueTypeExcell.Double)
            {
                cell.SetCellValue(double.Parse(data.Value));
            }
            if (data.ValueType == ValueTypeExcell.String)
            {
                cell.SetCellValue(data.Value);
            }
            if (data.ValueType == ValueTypeExcell.DateTime)
            {
                cell.SetCellValue(Convert.ToDateTime(data.Value));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellAddress"></param>
        /// <param name="addressFormat"></param>
        /// <returns></returns>
        public static (int ColumnIndex, int RowIndex, bool isSuccses) GetCellAddress(string cellAddress, AddressFormat addressFormat)
        {
            if (addressFormat == AddressFormat.A1)
            {
                return GetCellAddressFromA1(cellAddress);
            }
            if (addressFormat == AddressFormat.R1C1)
            {
                return GetCellAddressFromR1C1(cellAddress);
            }
            return GetCellAddressFromA1(cellAddress);
        }
        private static (int ColumnIndex, int RowIndex, bool isSuccses) GetCellAddressFromR1C1(string cellAddressFormatR1C1)
        {
            int columnIndex = 0;
            int rowIndex = 0;
            int index = 0;
            if (index >= cellAddressFormatR1C1.Length || cellAddressFormatR1C1[index] != 'R' & cellAddressFormatR1C1[index] != 'r')
            {
                return (columnIndex, rowIndex, false);
            }
            string rowName = string.Empty;
            index++;
            while (char.IsDigit(cellAddressFormatR1C1[index]) && index < cellAddressFormatR1C1.Length)
            {
                rowName += cellAddressFormatR1C1[index];
                index++;
            }
            if (index >= cellAddressFormatR1C1.Length || cellAddressFormatR1C1[index] != 'C' & cellAddressFormatR1C1[index] != 'c')
            {
                return (columnIndex, rowIndex, false);
            }
            string columnName = string.Empty;
            index++;
            while (index < cellAddressFormatR1C1.Length && char.IsDigit(cellAddressFormatR1C1[index]))
            {
                columnName += cellAddressFormatR1C1[index];
                index++;
            }
            if (string.IsNullOrEmpty(columnName) || string.IsNullOrEmpty(rowName))
            {
                return (columnIndex, rowIndex, false);
            }
            columnIndex = int.Parse(columnName) - 1;
            rowIndex = int.Parse(rowName) - 1;
            if (columnIndex <= 0 || rowIndex <= 0)
            {
                return (columnIndex, rowIndex, false);
            }
            return (columnIndex, rowIndex, true);
        }
        private static (int ColumnIndex, int RowIndex, bool isSuccses) GetCellAddressFromA1(string cellAddressFormatA1)
        {
            int columnIndex = 0;
            int rowIndex = 0;
            int firstNumIndex = 0;
            for (; firstNumIndex < cellAddressFormatA1.Length; firstNumIndex++)
            {
                if (Char.IsNumber(cellAddressFormatA1[firstNumIndex]))
                {
                    break;
                }
            }
            var arrayChar = cellAddressFormatA1.Substring(0, firstNumIndex).ToArray();
            Array.Reverse(arrayChar);
            string columnIndexStr = new string(arrayChar);
            for (int j = 0; j < columnIndexStr.Length; j++)
            {
                int charIndex = GetNumInAlphabet(columnIndexStr[j]);
                if (charIndex < 1 || charIndex > 26)
                {
                    return (columnIndex, rowIndex, false);
                }
                columnIndex += charIndex * (int)Math.Pow(26, j);
            }
            string rowIndexStr = cellAddressFormatA1.Substring(firstNumIndex, cellAddressFormatA1.Length - firstNumIndex);
            if (!int.TryParse(rowIndexStr, out rowIndex))
            {
                return (columnIndex, rowIndex, false);
            }
            if (columnIndex < 1 || rowIndex < 1)
            {
                return (columnIndex, rowIndex, false);
            }
            return (columnIndex - 1, rowIndex - 1, true);
        }

        static int GetNumInAlphabet(char ch)
        {
            return char.ToUpper(ch) - 64;
        }
        /// <summary>
        /// Возваращает экземпляр для записи текста в эксель 
        /// </summary>
        /// <param name="value"> текст для записи</param>
        /// <param name="cellAddressFormatA1"> дарес записи в формате A1</param>
        /// <returns></returns>
        static DataForExcell GetDataStringA1(string value, string cellAddressFormatA1)
        {
            return new DataForExcell(value, ValueTypeExcell.String, cellAddressFormatA1);
        }
        /// <summary>
        /// Возваращает экземпляр для записи текста в эксель
        /// </summary>
        /// <param name="value">число в форме строки для записи</param>
        /// <param name="cellAddressFormatA1">дарес записи в формате A1</param>
        /// <returns></returns>
        static DataForExcell GetDataDoubleA1(string value, string cellAddressFormatA1)
        {
            return new DataForExcell(value, ValueTypeExcell.Double, cellAddressFormatA1);
        }
    }
    [Serializable]
    /// <summary>
    /// Данные для внесения в документ
    /// </summary>
    public class DataForExcell
    {
        public DataForExcell()
        {

        }
        public DataForExcell(string value, ValueTypeExcell valueType, string cellAddress, AddressFormat addressFormat = AddressFormat.A1, int sheetNumber = 1, Action actionForExeption = null, string fileName = "")
        {
            if (addressFormat == AddressFormat.A1)
            {
                if (SetCellAddressA1(cellAddress) != ToExcellResult.Success)
                {
                    actionForExeption?.Invoke();
                }
            }
            if (addressFormat == AddressFormat.R1C1)
            {
                if (SetCellAddressR1C1(cellAddress) != ToExcellResult.Success)
                {
                    actionForExeption?.Invoke();
                }
            }
            Value = value;
            ValueType = valueType;
            SheetNumber = sheetNumber;
            FileName = fileName;
        }
        public ValueTypeExcell ValueType;
        public string Value { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string CellAddressA1
        {
            get { return cellAddressA1; }
            private set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    var columnRowIndex = ToExcell.GetCellAddress(value, AddressFormat.A1);
                    if (columnRowIndex.isSuccses == false)
                    {
                        throw new Exception("Введенное значение не адресс в формате A1");
                    }
                    RowIndex = columnRowIndex.RowIndex;
                    ColumnIndex = columnRowIndex.ColumnIndex;
                    cellAddressA1 = value;
                    cellAddressR1C1 = "R" + (columnRowIndex.RowIndex + 1) + "C" + (columnRowIndex.ColumnIndex + 1);
                }
            }
        }
        private string cellAddressA1;
        public string CellAddressR1C1
        {
            get { return cellAddressR1C1; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    var columnRowIndex = ToExcell.GetCellAddress(value, AddressFormat.R1C1);
                    if (columnRowIndex.isSuccses == false)
                    {
                        throw new Exception("Введенное значение не адресс в формате R1C1");
                    }
                    RowIndex = columnRowIndex.RowIndex;
                    ColumnIndex = columnRowIndex.ColumnIndex;
                    cellAddressR1C1 = value;
                    cellAddressA1 = GetColumnName(ColumnIndex) + (RowIndex + 1);
                }
            }
        }
        private string cellAddressR1C1;
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }
        public int SheetNumber
        {
            get { return SheetIndex + 1; }
            set
            {
                SheetIndex = value - 1;
                if (SheetIndex < 0)
                {
                    throw new Exception("Введенное значение не номер листа");
                }

            }
        }
        public int SheetIndex { get; private set; } = 0;
        public ToExcellResult SetCellAddressA1(string cellAddressA1)
        {
            try
            {
                CellAddressA1 = cellAddressA1;
            }
            catch (Exception exp)
            {
                return ToExcellResult.ParamError;
            }
            return ToExcellResult.Success;
        }
        public ToExcellResult SetCellAddressR1C1(string cellAddressR1C1)
        {
            try
            {
                CellAddressR1C1 = cellAddressR1C1;
            }
            catch (Exception exp)
            {
                return ToExcellResult.ParamError;
            }
            return ToExcellResult.Success;
        }
        private string GetColumnName(int columnIndex)
        {
            if (columnIndex < 26)
            {
                return "" + (char)(columnIndex + 1 + 64);
            }
            string columnName = string.Empty;
            while (columnIndex > 26)
            {
                columnName += (char)(columnIndex - (int)(columnIndex / 26) * 26 + 1 + 64);

                columnIndex = (int)(columnIndex / 26);
            }
            columnName += (char)(columnIndex + 64);
            char[] columNameRevers = columnName.ToArray();
            Array.Reverse(columNameRevers);
            return new string(columNameRevers);
        }
    }
}
