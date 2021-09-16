using DrillingBuildLibrary.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows;

namespace DrillingBuildLibrary
{
    public class NPOIHelper
    {
        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <returns>返回datatable</returns>
        public static DataSet ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataSet dataSet = null;
            FileStream fs = null;
            IWorkbook workbook = null;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        dataSet = new DataSet();

                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            ISheet sheet = workbook.GetSheetAt(i);
                            int startRow = 0;

                            //读取第一个sheet，当然也可以循环读取每个sheet
                            DataTable dataTable = ReadSheetData(isColumnName, sheet, ref startRow);
                            dataTable.TableName = sheet.SheetName;
                            dataSet.Tables.Add(dataTable);
                        }
                    }
                }
                return dataSet;
            }
            catch (Exception )
            {
                if (fs != null)
                {
                    fs.Close();
                }
               
                throw new Exception("读取excel失败,请关闭打开的excel文件！");
            }
        }

        private static DataTable ReadSheetData(bool isColumnName, ISheet sheet, ref int startRow)
        {
            DataTable dataTable = new DataTable();
            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;//总行数
                if (rowCount > 0)
                {
                    IRow firstRow = sheet.GetRow(0);//第一行
                    int cellCount = firstRow.LastCellNum;//列数
                    DataColumn column;
                    ICell cell;
                    
                    //构建datatable的列
                    if (isColumnName)
                    {
                        startRow = 1;//如果第一行是列名，则从第二行开始读取
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            cell = firstRow.GetCell(i);
                            IRichTextString richTextString = cell.RichStringCellValue;
                           
                            if (cell != null)
                            {
                                if (cell.StringCellValue != null)
                                {
                                    column = new DataColumn(cell.StringCellValue);
                                    if (dataTable.Columns.Contains(column.ColumnName))
                                    {
                                        dataTable.Columns.Add(column+"_1");

                                    }
                                    else
                                    {
                                        dataTable.Columns.Add(column);

                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            column = new DataColumn("column" + (i + 1));
                            dataTable.Columns.Add(column);
                        }
                    }

                    //填充行
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        DataRow dataRow = dataTable.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            cell = row.GetCell(j);
                            if (cell == null)
                            {
                                dataRow[j] = "";
                            }
                            else
                            {
                                //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                switch (cell.CellType)
                                {
                                    case CellType.Blank:
                                        dataRow[j] = "";
                                        break;
                                    case CellType.Numeric:
                                    case CellType.Formula:

                                        short format = cell.CellStyle.DataFormat;
                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                            dataRow[j] = cell.DateCellValue;
                                        else
                                            dataRow[j] = cell.NumericCellValue;
                                        break;
                                    case CellType.String:
                                        string str = "";
                                        XSSFRichTextString richTextString = cell.RichStringCellValue as XSSFRichTextString;
                                        var r= richTextString.GetCTRst().r;
                                        
                                        foreach (var  cT_RElt in r)
                                        {
                                          
                                            if (cT_RElt.rPr==null)
                                            {
                                                str += cT_RElt.t;
                                            }
                                            else
                                            {
                                                if (cT_RElt.rPr.vertAlign.val== NPOI.OpenXmlFormats.Spreadsheet.ST_VerticalAlignRun.subscript)
                                                {
                                                    foreach (var item in cT_RElt.t)
                                                    {
                                                        str += GetSuberChar(item.ToString());

                                                    }

                                                }
                                                else if (cT_RElt.rPr.vertAlign.val == NPOI.OpenXmlFormats.Spreadsheet.ST_VerticalAlignRun.superscript)
                                                {
                                                    foreach (var item in cT_RElt.t)
                                                    {
                                                        str += GetSuperChar(item.ToString());

                                                    }
                                                }
                                            }
                                        }

                                        dataRow[j] = str;
                                        break;
                                }
                            }
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            return dataTable;
        }

        /// <summary>
        /// 上标
        /// </summary>
        /// <returns></returns>
        private static string GetSuperChar(string str)
        {
            switch (str)
            {
                case "0":
                    return "\x2070";
                case "1":
                    return "\x00B9";
                case "2":
                    return "\x00B2";
                case "3":
                    return "\x00B3";
                case "4":
                    return "\x2074";
                case "5":
                    return "\x2075";
                case "6":
                    return "\x2076";
                case "7":
                    return "\x2077";
                case "8":
                    return "\x2078";
                case "9":
                    return "\x2079";
                case "+":
                    return "\x207A";
                case "-":
                    return "\x207B";
                case "=":
                    return "\x207C";
                case "(":
                    return "\x207D";
                case ")":
                    return "\x207E";
                case "i":
                    return "\x2071";
                case "n":
                    return "\x207F";
                default:

                    return $"<sup>{str}</sup>";
            }
        }
        /// <summary>
        /// 下标
        /// </summary>
        /// <returns></returns>
        private static string GetSuberChar(string str)
        {
            switch (str)
            {
                case "0":
                    return "\x2080";
                case "1":
                    return "\x2081";
                case "2":
                    return "\x2082";
                case "3":
                    return "\x2083";
                case "4":
                    return "\x2084";
                case "5":
                    return "\x2085";
                case "6":
                    return "\x2086";
                case "7":
                    return "\x2087";
                case "8":
                    return "\x2088";
                case "9":
                    return "\x2089";
                case "+":
                    return "\x208A";
                case "-":
                    return "\x208B";
                case "=":
                    return "\x208C";
                case "(":
                    return "\x208D";
                case ")":
                    return "\x208E";
                case "a":
                    return "\x2090";
                case "e":
                    return "\x2091";
                case "o":
                    return "\x2092";
                case "i":
                    return "\x1D62";
                case "r":
                    return "\x1D63";
                case "u":
                    return "\x1D64";
                case "v":
                    return "\x1D65";
                case "x":
                    return "\x2093";
                case "β":
                    return "\x1D66";
                case "γ":
                    return "\x1D67";
                case "χ":
                    return "\x1D6A";
                case "ψ":
                    return "\x1D69";
                case "h":
                    return "ₕ";
                case "k":
                    return "ₖ";
                case "l":
                    return "ₗ";
                case "m":
                    return "ₘ";
                case "n":
                    return "ₙ";
                case "p":
                    return "ₚ";
                case "s":
                    return "ₛ";
                case "t":
                    return "ₜ";
                default:
                    return $"<sub>{str}</sub>";
            }
        }
        public static bool DatasetToExcel(DataSet dataset, string strFile)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            try
            {
                if (dataset.Tables.Count>0)
                {
                    // 2007版本
                    if (strFile.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook();
                    // 2003版本
                    else if (strFile.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook();
                    for (int tableIndex = 0; tableIndex < dataset.Tables.Count; tableIndex++)
                    {
                        DataTable dataTable = dataset.Tables[tableIndex];
                        ISheet sheet = null;
                        if (dataTable.TableName!="")
                        {
                            sheet = workbook.CreateSheet(dataTable.TableName);

                        }
                        else
                        {
                            sheet = workbook.CreateSheet("Sheet"+ (tableIndex+1));//
                        }

                        IRow row = null;
                        ICell cell = null;
                        int rowCount = dataTable.Rows.Count;//行数
                        int columnCount = dataTable.Columns.Count;//列数

                        //设置列头
                        row = sheet.CreateRow(0);//excel第一行设为列头
                        for (int c = 0; c < columnCount; c++)
                        {
                            cell = row.CreateCell(c);
                            cell.SetCellValue(dataTable.Columns[c].ColumnName);
                        }

                        //设置每行每列的单元格,
                        for (int i = 0; i < rowCount; i++)
                        {
                            row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < columnCount; j++)
                            {
                                cell = row.CreateCell(j);//excel第二行开始写入数据
                                cell.SetCellValue(dataTable.Rows[i][j].ToString());
                            }
                        }
                    }

                    using (fs = File.OpenWrite(strFile))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据
                        result = true;
                    }
                }
            }
            catch (Exception)
            {
                fs?.Dispose();
                throw new Exception("读取excel失败！");
            }
            return result;
        }

        /// <summary>
        /// 写入excel
        /// </summary>
        /// <param name="dt">datatable</param>
        /// <param name="strFile">strFile</param>
        /// <returns></returns>
        public static bool DataTableToExcel(DataTable dt, string strFile)
        {

            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            // 2007版本
            if (strFile.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook();
            // 2003版本
            else if (strFile.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook();
            bool result = false;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    //设置列头
                    row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (fs = File.OpenWrite(strFile))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception )
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }
        }
        /// <summary>
        /// 写入模板表excel
        /// </summary>
        /// <param name="dt">datatable</param>
        /// <param name="strFile">strFile</param>
        /// <returns></returns>
        public static bool DataModelTableToExcel(DataTable dt, string strFile, Dictionary<string, string[]> dropDownLists)
        {
            bool result = false;
            IWorkbook workbook = null;
            // 2007版本
            if (strFile.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook();
            // 2003版本
            else if (strFile.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook();

            //
            ICellStyle heardStyle = workbook.CreateCellStyle();//声明style1对象，设置Excel表格的样式
            ICellStyle bodyStyle = workbook.CreateCellStyle();

            IFont heardFont = workbook.CreateFont();
            //font.Color = IndexedColors.Red.Index;
            heardFont.FontHeightInPoints = 16;
            heardFont.FontName = "等线";
            heardFont.IsBold = true;

            IFont bodyFont = workbook.CreateFont();
            //font.Color = IndexedColors.Red.Index;
            bodyFont.FontHeightInPoints = 11;
            bodyFont.FontName = "等线";
            bodyFont.IsBold = true;

            //设置表头样式
            heardStyle.SetFont(heardFont);
            //设置表头对齐方式
            heardStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            heardStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

            //heardStyle.BorderLeft = BorderStyle.Thick;
            //heardStyle.BorderRight = BorderStyle.Thick;
            //heardStyle.BorderTop = BorderStyle.Thick;
            //heardStyle.BorderBottom = BorderStyle.Thick;

            //设置主体样式
            bodyStyle.SetFont(bodyFont);
            //设置主体对齐方式
            bodyStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Justify;//两端自动对齐（自动换行）
            bodyStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

            //bodyStyle.BorderLeft = BorderStyle.Thick;
            //bodyStyle.BorderRight = BorderStyle.Thick;
            //bodyStyle.BorderTop = BorderStyle.Thick;
            //bodyStyle.BorderBottom= BorderStyle.Thick;

            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            
            string[] sArray = dt.TableName.Split('-');
            try
            {
                if (dt != null )
                {
                    sheet = workbook.CreateSheet( $"{sArray[1]}" );//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    if (rowCount > 1)
                        ;

                    //设置模板名称
                    CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnCount-1);
                    sheet.AddMergedRegion(region);
                    row = sheet.CreateRow(0);
                    ICell heardCell = row.CreateCell(0);
                    heardCell.SetCellValue(sArray[1]);
                    
                    //模板样式
                    heardCell.CellStyle = heardStyle;

                    //设置列头
                    row = sheet.CreateRow(1);//excel第二行行设为列头

                    //下拉框个数
                    int dropDownListCout = 1;
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);

                        //拆分
                        string[] columnNameArray = dt.Columns[c].ColumnName.Split('-');
                        if (dropDownLists.ContainsKey(columnNameArray[0]))
                        {
                            string[] dropDownListValues;
                            //返回对应的
                            dropDownLists.TryGetValue(columnNameArray[0], out dropDownListValues);
                            SetCellDropdownList(workbook, sheet, columnNameArray[1], c, c, dropDownListValues, dropDownListCout);
                            dropDownListCout++;
                        }

                        
                        //向表中增加值
                        cell.SetCellValue(columnNameArray[1]);
                       


                        //设置主体样式
                        cell.CellStyle = bodyStyle;
                        sheet.AutoSizeColumn(c,true);

                    }
                    using (fs = File.OpenWrite(strFile))
                    {

                        workbook.Write(fs);//向打开的这个xls文件中写入数据

                        result = true;
                    }


                }
                return result;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Dispose();
                    fs.Close();
                }
                return false;
            }
        }


        /// <summary>
        /// 下拉框
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheet"></param>
        /// <param name="name"></param>
        /// <param name="firstcol"></param>
        /// <param name="lastcol"></param>
        /// <param name="vals"></param>
        /// <param name="sheetindex"></param>
        public static void SetCellDropdownList(IWorkbook workbook, ISheet sheet, string name, int firstcol, int lastcol, string[] vals, int sheetindex)
        {
            //先创建一个Sheet专门用于存储下拉项的值
            ISheet sheet2 = workbook.CreateSheet(name);
            //隐藏
            workbook.SetSheetHidden(sheetindex, SheetState.VeryHidden);
            int index = 0;
            foreach (var item in vals)
            {
                sheet2.CreateRow(index).CreateCell(0).SetCellValue(item);
                index++;
            }
            //创建的下拉项的区域：
            var rangeName = name+"Rang";
            IName range = workbook.CreateName();
            range.RefersToFormula = name + "!$A$1:$A$" + index;
            range.NameName = rangeName;
            CellRangeAddressList regions = new CellRangeAddressList(2, 65535, firstcol, lastcol);//约束范围

            ISheet sheet1 = workbook.GetSheetAt(0);//获得第一个工作表
            XSSFDataValidationHelper helper = new XSSFDataValidationHelper((XSSFSheet)sheet1);//获得一个数据验证Helper
            IDataValidation validation = helper.CreateValidation(helper.CreateFormulaListConstraint($"{rangeName}"), regions);//创建一个特定约束范围内的公式列表约束（即第一节里说的"自定义"方式）
            validation.CreateErrorBox("错误", "请按右侧下拉箭头选择!");//不符合约束时的提示
            validation.ShowErrorBox = true;//显示上面提示 = True
            sheet1.AddValidationData(validation);//添加进去
            sheet1.ForceFormulaRecalculation = true;
        }
        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string file)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);

                //表头
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// Datable导出成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file">导出路径(包括文件名与扩展名)</param>
        public static void TableToExcel(DataTable dt, string file)
        {
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:
                    return null;
                case CellType.Boolean: //BOOLEAN:
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:
                    return cell.NumericCellValue;
                case CellType.String: //STRING:
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:
                default:
                    return "=" + cell.CellFormula;

            }
        }
    }
}
