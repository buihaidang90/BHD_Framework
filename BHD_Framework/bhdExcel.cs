using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.Globalization;
using System.Drawing;

namespace BHD_Framework
{
    public class bhdExcel
    {
        //private string _CheckSymbol = "☒";
        //private string _UnCheckSymbol = "☐";
        //public bhdExcel(DataTable InputTable)
        //{
        //    Init_Config();
        //    this.dtbMain = InputTable;
        //    cols_Set_All();
        //}
        

        ////public 

        ///// <summary>
        ///// Lấy chỉ số của sheet trong file excel
        ///// </summary>
        ///// <param name="ExcelFilePath">đường dẫn file excel</param>
        ///// <param name="SheetName">tên sheet</param>
        ///// <returns></returns>
        //public int SheetIndex(string ExcelFilePath, string SheetName)
        //{
        //    int ind = -1;
        //    try
        //    {
        //        using (FileStream file = new FileStream(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //        {
        //            IWorkbook workbook = WorkbookFactory.Create(file);
        //            ind = workbook.GetSheetIndex(SheetName);
        //            return ind;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("Your file excel is invalid", ex);
        //    }
        //}

        ///// <summary>
        ///// Lấy tên tất cả các sheet trong file excel
        ///// </summary>
        ///// <param name="ExcelFilePath">đường dẫn file excel</param>
        ///// <returns></returns>
        //public string[] SheetNames(string ExcelFilePath)
        //{
        //    try
        //    {
        //        using (FileStream file = new FileStream(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //        {
        //            IWorkbook workbook = WorkbookFactory.Create(file);
        //            string[] s = new string[workbook.NumberOfSheets];
        //            for (int i = 0; i < workbook.NumberOfSheets; i++)
        //            {
        //                s[i] = workbook.GetSheetName(i);
        //            }
        //            return s;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message.ToLower().Contains("invalid uri"))
        //            throw new Exception("Check excel file and remove all invalid hyperlink!", ex);
        //        return null;
        //    }
        //}

        ///// <summary>
        ///// Import file excel bằng thư viện NPOI, hỗ trợ định dạng xls và xlsx
        ///// </summary>
        ///// <param name="ExcelFilePath">đường dẫn file excel</param>
        ///// <param name="SheetName">tên sheet cần xử lý (ko truyền thì mặc định lấy sheet đầu tiên)</param>
        ///// <param name="HeaderLineNumber">tổng số dòng header</param>
        ///// <returns></returns>
        //public DataTable Import(string ExcelFilePath, string SheetName, int HeaderLineNumber)
        //{
        //    try
        //    {
        //        if (HeaderLineNumber < 0) HeaderLineNumber = 0;
        //        int iHeaderRow = HeaderLineNumber - 1;
        //        int iBeginRow = HeaderLineNumber;
        //        int iEndCol = 0;
        //        IRow r; ICell c;
        //        DataTable dt = new DataTable();
        //        #region XLS - Excel 2003 and below
        //        if (ExcelFilePath.ToUpper().EndsWith(".XLS"))
        //        {
        //            using (FileStream fs = new FileStream(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //            {
        //                HSSFWorkbook workbook = new HSSFWorkbook(fs);
        //                ISheet s;
        //                if (SheetName.Trim() == string.Empty) s = workbook.GetSheetAt(0);
        //                else if (workbook.GetSheetIndex(SheetName) >= 0) s = workbook.GetSheet(SheetName);
        //                else throw new Exception("The sheet name does not exist !");
        //                System.Collections.IEnumerator rows = s.GetRowEnumerator();
        //                while (rows.MoveNext())
        //                {
        //                    r = rows.Current as HSSFRow;
        //                    if (r.RowNum == iHeaderRow)
        //                    {// dòng header = dòng tiêu đề của file excel
        //                        iEndCol = 0;
        //                        for (int i = 0; i < r.LastCellNum; i++)
        //                        {
        //                            c = r.GetCell(i);
        //                            if (c == null) break;
        //                            else if (c.ToString().Trim() == string.Empty) break;
        //                            else
        //                            {
        //                                // cắt bỏ ký tự [Alt+Enter] trong 1 ô excel rồi đưa 2 dòng lấy được vào mảng, sau đó lấy dòng đầu tiên làm tiêu đề cột
        //                                try { dt.Columns.Add(c.ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim()); }
        //                                catch { dt.Columns.Add("Column " + (i + 1).ToString().Trim()); }
        //                            }
        //                            iEndCol++;
        //                        }// end for
        //                    }
        //                    if (r.RowNum < iBeginRow) continue;
        //                    //=Data=================================
        //                    DataRow dr = dt.NewRow();
        //                    for (int i = 0; i < iEndCol; i++)
        //                    {
        //                        c = r.GetCell(i);
        //                        //if (c == null) dr[i] = null;
        //                        //else dr[i] = c.ToString();
        //                        dr[i] = ParseNull(c);
        //                    }
        //                    dt.Rows.Add(dr);
        //                }// end while
        //            }// end using
        //        }
        //        #endregion
        //        #region XLSX - Excel 2007 and above
        //        if (ExcelFilePath.ToUpper().EndsWith(".XLSX"))
        //        {
        //            using (FileStream fs = new FileStream(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //            {
        //                XSSFWorkbook workbook = new XSSFWorkbook(fs);
        //                ISheet s;
        //                if (SheetName.Trim() == string.Empty) s = workbook.GetSheetAt(0);
        //                else if (workbook.GetSheetIndex(SheetName) >= 0) s = workbook.GetSheet(SheetName);
        //                else throw new Exception("The sheet name does not exist !");
        //                System.Collections.IEnumerator rows = s.GetRowEnumerator();
        //                while (rows.MoveNext())
        //                {
        //                    r = rows.Current as HSSFRow;
        //                    if (r.RowNum == iHeaderRow)
        //                    {// dòng header = dòng tiêu đề của file excel
        //                        iEndCol = -1;
        //                        for (int i = 0; i < r.LastCellNum; i++)
        //                        {
        //                            c = r.GetCell(i);
        //                            if (c == null) break;
        //                            else if (c.ToString().Trim() == string.Empty) break;
        //                            else
        //                            {
        //                                // cắt bỏ ký tự [Alt+Enter] trong 1 ô excel rồi đưa 2 dòng lấy được vào mảng, sau đó lấy dòng đầu tiên làm tiêu đề cột
        //                                try { dt.Columns.Add(c.ToString().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim()); }
        //                                catch { dt.Columns.Add("Column " + (i + 1).ToString().Trim()); }
        //                            }
        //                            iEndCol++;
        //                        }// end for
        //                    }
        //                    if (r.RowNum < iBeginRow) continue;
        //                    //=Data=================================
        //                    DataRow dr = dt.NewRow();
        //                    for (int i = 0; i < iEndCol; i++)
        //                    {
        //                        c = r.GetCell(i);
        //                        //if (c == null) dr[i] = null;
        //                        //else dr[i] = c.ToString();
        //                        dr[i] = ParseNull(c);
        //                    }
        //                    dt.Rows.Add(dr);
        //                }// end while
        //            }// end using
        //        }
        //        #endregion
        //        return FixDataImport(dt);
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message.ToLower().Contains("invalid uri"))
        //            throw new Exception("Check excel file and remove all invalid hyperlink!", ex);
        //        return null;
        //    }
        //}

        ///// <summary>
        ///// Import file excel bằng thư viện NPOI, hỗ trợ định dạng xls và xlsx
        ///// </summary>
        ///// <param name="ExcelFilePath">đường dẫn file excel</param>
        ///// <param name="SheetName">tên sheet cần xử lý (ko truyền thì mặc định lấy sheet đầu tiên)</param>
        ///// <returns></returns>
        //public DataTable Import(string ExcelFilePath, string SheetName)
        //{
        //    return this.Import(ExcelFilePath, SheetName, 1);
        //}

        ///// <summary>
        ///// Import file excel bằng thư viện NPOI, hỗ trợ định dạng xls và xlsx
        ///// </summary>
        ///// <param name="ExcelFilePath">đường dẫn file excel</param>
        ///// <returns></returns>
        //public DataTable Import(string ExcelFilePath)
        //{
        //    return this.Import(ExcelFilePath, "", 1);
        //}

        //private DataTable FixDataImport(DataTable dt)
        //{
        //    DataTable dtFixed = dt.Clone();
        //    // chuyển toàn bộ dữ liệu về kiểu chuỗi
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //        dtFixed.Columns[i].DataType = typeof(string);

        //    //DataTable.Load() = copy dữ liệu từ datareader; DataTable.CreateDataReader() = tạo bảng dữ liệu để đổ vào datatable
        //    dtFixed.Load(dt.CreateDataReader());

        //    //for (int i = 0; i < dtFixed.Rows.Count; i++)
        //    //    for (int j = 0; j < dtFixed.Columns.Count; j++) 
        //    //        dtFixed.Rows[i][j] = ParseNull(dtFixed.Rows[i][j]);

        //    for (int i = dtFixed.Rows.Count - 1; i >= 0; i--)
        //    {
        //        bool isRowNull = false;
        //        for (int j = 0; j < dtFixed.Columns.Count; j++)
        //        {
        //            if (dtFixed.Rows[i][j].ToString() != "")
        //            {
        //                isRowNull = true; break;
        //            }
        //        }
        //        if (!isRowNull) dtFixed.Rows.RemoveAt(i);
        //    }
        //    return dtFixed;
        //}

        //private string ParseNull(object InputObject)
        //{
        //    if (InputObject == null) return string.Empty;
        //    else return InputObject.ToString();
        //}


        //private void Init_Config()
        //{
        //    try
        //    {
        //        //mksFrameworkConfig fc = new mksFrameworkConfig();
        //        //if (fc.Group_Check("mksExcel")) 
        //        //{
        //        //    if (fc.Key_Check("__Use_ActiveFont"))
        //        //        __Use_ActiveFont = fc.Get_Bool("__Use_ActiveFont");
        //        //    if (fc.Key_Check("__HeaderColor"))
        //        //        Caption_Color_Back = Color.FromArgb(fc.Get_Int("__HeaderColor"));
        //        //    if (fc.Key_Check("Caption_Color_Back"))
        //        //        Caption_Color_Back = Color.FromArgb(fc.Get_Int("Caption_Color_Back"));
        //        //    if (fc.Key_Check("Caption_Color_Fore"))
        //        //        Caption_Color_Fore = Color.FromArgb(fc.Get_Int("Caption_Color_Fore"));
        //        //    if (__Use_ActiveFont)
        //        //    {
        //        //        string str = System.Windows.Forms.Application.OpenForms[0].Controls[0].Font.Name;
        //        //        if (str != "Microsoft Sans Serif" && str != "Tahoma")
        //        //        {
        //        //            Default_Font_Name_ = str;
        //        //            DeFont = new Font(Default_Font_Name_, 11);
        //        //            DeFontB = new Font(Default_Font_Name_, 11, FontStyle.Bold);
        //        //        }
        //        //    }
        //        //    if (fc.Key_Check("Default_Font_Name"))
        //        //        Default_Font_Name_ = fc.Get_String("Default_Font_Name");
        //        //    if (fc.Key_Check("Default_Font_Size"))
        //        //        Default_Font_Size_ = Convert.ToSingle(fc.Get_Double("Default_Font_Size"));
        //        //    if (fc.Key_Check("Default_Font_Name") || fc.Key_Check("Default_Font_Size")) {
        //        //        DeFont = new Font(Default_Font_Name_, Default_Font_Size_);
        //        //        DeFontB = new Font(Default_Font_Name_, Default_Font_Size_, FontStyle.Bold);
        //        //    }
        //        //}
        //    }
        //    catch { }
        //}

        //private void cols_Set_All()
        //{
        //    cols_Type = new List<string>();
        //    cols_Format = new List<string>();
        //    cols_AutoWidth = new List<int>();
        //    cols_CustomWidth = new List<int>();
        //    cols_Visible = new List<int>();
        //    if (dtbMain != null)
        //    {
        //        for (int i = 0; i < dtbMain.Columns.Count; i++)
        //        {
        //            cols_Type.Add(dtbMain.Columns[i].DataType.FullName);
        //            if (cols_Type[i] == tString || cols_Type[i] == tDateTime)
        //                cols_Format.Add("@");
        //            else
        //                cols_Format.Add("");
        //            cols_Visible.Add(i);
        //            cols_AutoWidth.Add(0); cols_CustomWidth.Add(0);
        //        }
        //    }
        //}

        //private void cs_Copy(ICellStyle FromCS, ICellStyle ToCS)
        //{
        //    ToCS.CloneStyleFrom(FromCS);
        //}
        //private void cs_SetBackColor_XLS(ICellStyle InputICellStyle, Color InputColor)
        //{
        //    HSSFCellStyle cs = (HSSFCellStyle)InputICellStyle;
        //    cs.FillForegroundColor = cl_Find_XLS(InputColor);
        //    cs.FillPattern = FillPattern.SolidForeground;
        //}
        //private void cs_SetBackColor_XLSX(ICellStyle InputICellStyle, Color InputColor)
        //{
        //    XSSFCellStyle cs = (XSSFCellStyle)InputICellStyle;
        //    //cs.FillForegroundXSSFColor = cl_Find_XLSX(InputColor); 
        //    cs.SetFillForegroundColor(cl_Find_XLSX(InputColor));
        //    cs.FillPattern = FillPattern.SolidForeground;
        //}
        //private void cs_SetForeColor_XLS(ICellStyle InputICellStyle, Color InputColor, IWorkbook InputWorkbook)
        //{
        //    HSSFCellStyle cs = (HSSFCellStyle)InputICellStyle;
        //    cs.GetFont((IWorkbook)InputWorkbook).Color = cl_Find_XLS(InputColor);
        //}
        //private void cs_SetForeColor_XLSX(ICellStyle InputICellStyle, Color InputColor, IWorkbook InputWorkbook)
        //{
        //    ICellStyle cs = (ICellStyle)InputICellStyle;
        //    //XSSFFont ft = (XSSFFont)cs.GetFont((IWorkbook)InputWorkbook);
        //    //ft. SetColor(cl_Get_XLSX( InputColor));
        //    //cs.SetFont(ft);
        //    XSSFFont ft = ((XSSFFont)cs.GetFont((IWorkbook)InputWorkbook));
        //    ft.SetColor(cl_Find_XLSX(InputColor));
        //    cs.SetFont(ft);
        //}

        //private bool UseTreeNode_XL = false;//use for grid
        //private char _TreeNode = ' ';
        //private void UseTreeNode(bool bActivate, char cTreeNode)
        //{
        //    UseTreeNode_XL = bActivate;
        //    this._TreeNode = cTreeNode;
        //}
        
        //private List<string> MergedRange_Check = new List<string>();

        //#region xlHeader
        //public enum HAlign { Left, Center, Right }
        //public enum VAlign { Top, Center, Bottom }
        //public struct xlHeader
        //{
        //    public string Text;
        //    public HAlign HorizontalAlignment;
        //    public VAlign VerticalAlignment;
        //    public Color ForeColor;
        //    public string FontName;
        //    public int FontSize;
        //    public bool FontBold;

        //    public xlHeader(string InputText)
        //    {
        //        Text = InputText;
        //        HorizontalAlignment = HAlign.Left;
        //        VerticalAlignment = VAlign.Center;
        //        ForeColor = Color.Black;
        //        FontName = TimesNewRoman;
        //        FontSize = 1923; //Memory
        //        FontBold = true;
        //    }
        //    public xlHeader(string InputText, int HeaderLevel)
        //    {
        //        Text = InputText;
        //        switch (HeaderLevel)
        //        {
        //            case 0: //Empty Line
        //                HorizontalAlignment = HAlign.Center;
        //                VerticalAlignment = VAlign.Center;
        //                ForeColor = Color.Blue;
        //                FontName = "--[BHD]--";
        //                FontSize = 1923; //Memory
        //                FontBold = true;
        //                break;
        //            case 1: //Main Header Line
        //                HorizontalAlignment = HAlign.Center;
        //                VerticalAlignment = VAlign.Center;
        //                ForeColor = Color.Blue;
        //                FontName = TimesNewRoman;
        //                FontSize = 14;
        //                FontBold = true;
        //                break;
        //            case 2: //Left Normal Line
        //                HorizontalAlignment = HAlign.Left;
        //                VerticalAlignment = VAlign.Center;
        //                ForeColor = Color.Black;
        //                FontName = TimesNewRoman;
        //                FontSize = 11;
        //                FontBold = false;
        //                break;
        //            case 3: //Center Normal Line
        //                HorizontalAlignment = HAlign.Center;
        //                VerticalAlignment = VAlign.Center;
        //                ForeColor = Color.Black;
        //                FontName = TimesNewRoman;
        //                FontSize = 11;
        //                FontBold = false;
        //                break;
        //            case 4: //Right Normal Line
        //                HorizontalAlignment = HAlign.Right;
        //                VerticalAlignment = VAlign.Center;
        //                ForeColor = Color.Black;
        //                FontName = TimesNewRoman;
        //                FontSize = 11;
        //                FontBold = false;
        //                break;
        //            default: //Default Header Line = Left Normal Line
        //                HorizontalAlignment = HAlign.Left;
        //                VerticalAlignment = VAlign.Center;
        //                ForeColor = Color.Black;
        //                FontName = TimesNewRoman;
        //                FontSize = 11;
        //                FontBold = true;
        //                break;
        //        }
        //    }
        //}
        //#endregion

        //#region Formular struct
        //private struct Formular_Struct
        //{
        //    public int SheetIndex;
        //    public int RowIndex;
        //    public int ColIndex;
        //    public string ExcelFormular;
        //    public Formular_Struct(int _SheetIndex, int _RowIndex, int _ColIndex, string _ExcelFormular)
        //    {
        //        SheetIndex = _SheetIndex;
        //        RowIndex = _RowIndex;
        //        ColIndex = _ColIndex;
        //        ExcelFormular = _ExcelFormular;
        //    }
        //}

        //private List<Formular_Struct> Formular_Arr = new List<Formular_Struct>();
        ////public void Formular_Set(int _SheetIndex, int _RowIndex, int _ColIndex, string _ExcelFormular)
        ////{
        ////    Formular_Struct fs = new Formular_Struct(_SheetIndex, _RowIndex, _ColIndex, _ExcelFormular);
        ////    Formular_Arr.Add(fs);
        ////}
        //#endregion

        //#region Config
        //private DataTable dtbMain = null;

        //private bool __Use_ActiveFont = false; // change font-name
        //private Color Caption_Color_Back = Color.Azure;
        //private Color Caption_Color_Fore = Color.DarkBlue;

        //private bool UseCellBackColor = false; // use for grid
        //private bool UseCellForeColor = false; // use for grid
        //private bool UseCellFontStyle = false; // use for grid + XLS
        //private bool UseCellMergeing = false; // use for grid
        //private bool UseInvisibleRow = false; // use for grid

        //private bool UseCustomBoolean_XL = false;
        //private string Bool_True = @"T";
        //private string Bool_False = @"F";
        //private void UseCustomBoolean(string TrueText, string FalseText)
        //{
        //    if (TrueText != "" || FalseText != "")
        //    {
        //        Bool_True = TrueText;
        //        Bool_False = FalseText;
        //        UseCustomBoolean_XL = true;
        //    }
        //}

        //private const string TimesNewRoman = "Times New Roman";
        //private Font DeFont = new Font(TimesNewRoman, 11);
        //private Font DeFontB = new Font(TimesNewRoman, 11, FontStyle.Bold);
        //private string Default_Font_Name_ = TimesNewRoman;
        //[System.ComponentModel.DefaultValue(TimesNewRoman)] /*Microsoft Sans Serif*/
        //public string Default_Font_Name
        //{
        //    get { return Default_Font_Name_; }
        //    set
        //    {
        //        Default_Font_Name_ = value;
        //        DeFont = new Font(Default_Font_Name_, Default_Font_Size_);
        //        DeFontB = new Font(Default_Font_Name_, Default_Font_Size_, FontStyle.Bold);
        //    }
        //}

        //private float Default_Font_Size_ = 11;
        //[System.ComponentModel.DefaultValue(11)]
        //public float Default_Font_Size
        //{
        //    get { return Default_Font_Size_; }
        //    set
        //    {
        //        Default_Font_Size_ = value;
        //        DeFont = new Font(Default_Font_Name_, Default_Font_Size_);
        //        DeFontB = new Font(Default_Font_Name_, Default_Font_Size_, FontStyle.Bold);
        //    }
        //}

        //#endregion

        //#region Fullname of types
        //private string tString = typeof(string).FullName;
        //private string tInt16 = typeof(Int16).FullName;
        //private string tInt32 = typeof(Int32).FullName;
        //private string tInt64 = typeof(Int64).FullName;
        //private string tFloat = typeof(float).FullName;
        //private string tDouble = typeof(double).FullName;
        //private string tDecimal = typeof(decimal).FullName;
        //private string tBool = typeof(bool).FullName;
        //private string tDateTime = typeof(DateTime).FullName;
        //#endregion

        //private Graphics DeG = Graphics.FromImage(new Bitmap(500, 500));
        //private int tmp_Set_ColumnAutoWidth = 0;
        //public void Set_ColumnAutoWidth(int ColumnIndex, string TextInput)
        //{
        //    Set_ColumnAutoWidth(ColumnIndex, TextInput, false);
        //}
        //public void Set_ColumnAutoWidth(int ColumnIndex, string TextInput, bool IsBold)
        //{
        //    tmp_Set_ColumnAutoWidth = DeG.MeasureString(TextInput + "L", IsBold ? DeFontB : DeFont).ToSize().Width;
        //    if (cols_CustomWidth[ColumnIndex] < tmp_Set_ColumnAutoWidth)
        //        cols_CustomWidth[ColumnIndex] = tmp_Set_ColumnAutoWidth;
        //}

        //public void Set_ColumnWidth(int ColumnIndex, int WidthSize)
        //{
        //    cols_AutoWidth[ColumnIndex] = WidthSize;
        //}
        //public void ShowColumn(params int[] VisibleColumn_Index)
        //{
        //    foreach (int i in VisibleColumn_Index)
        //    {
        //        if (!cols_Visible.Contains(i))
        //            cols_Visible.Add(i);
        //    }
        //}
        //public void HideColumn(params int[] InvisibleColumn_Index)
        //{
        //    foreach (int i in InvisibleColumn_Index)
        //    {
        //        if (cols_Visible.Contains(i))
        //            cols_Visible.Remove(i);
        //    }
        //}

        //public List<xlHeader> HeaderArea = new List<xlHeader>();
        //public void Set_Header(string InputText, int HeaderLevel)
        //{
        //    xlHeader h = new xlHeader(InputText, HeaderLevel);
        //    h.FontName = Default_Font_Name_;
        //    HeaderArea.Add(h);
        //}


        //private void ft_Copy(IFont FromF, IFont ToF)
        //{
        //    ToF.Boldweight = FromF.Boldweight;
        //    ToF.Charset = FromF.Charset;
        //    ToF.Color = FromF.Color;
        //    ToF.FontHeight = FromF.FontHeight;
        //    ToF.FontHeightInPoints = FromF.FontHeightInPoints;
        //    ToF.FontName = FromF.FontName;
        //    ToF.IsItalic = FromF.IsItalic;
        //    ToF.IsStrikeout = FromF.IsStrikeout;
        //    ToF.TypeOffset = FromF.TypeOffset;
        //    ToF.Underline = FromF.Underline;
        //}

        //private XSSFFont ft_Tmp_XLSX = null;
        //private List<XSSFFont> ft_AllMain_XLSX = null;

        //private IFont ft_Tmp = null;
        //private List<IFont> ft_AllMain;
        //private IFont ft_Find(IFont iF, IWorkbook wb)
        //{
        //    foreach (IFont f in ft_AllMain)
        //    {
        //        if (f.Index == iF.Index) continue;
        //        if (f.Boldweight != iF.Boldweight) continue;
        //        if (f.Charset != iF.Charset) continue;
        //        if (f.Color != iF.Color) continue;
        //        if (f.FontHeight != iF.FontHeight) continue;
        //        if (f.FontHeightInPoints != iF.FontHeightInPoints) continue;
        //        if (f.FontName != iF.FontName) continue;
        //        if (f.IsItalic != iF.IsItalic) continue;
        //        if (f.IsStrikeout != iF.IsStrikeout) continue;
        //        if (f.TypeOffset != iF.TypeOffset) continue;
        //        if (f.Underline != iF.Underline) continue;
        //        return f;
        //    }
        //    IFont fNew = wb.CreateFont();
        //    fNew.Boldweight = iF.Boldweight;
        //    fNew.Charset = iF.Charset;
        //    fNew.Color = iF.Color;
        //    fNew.FontHeight = iF.FontHeight;
        //    fNew.FontHeightInPoints = iF.FontHeightInPoints;
        //    fNew.FontName = iF.FontName;
        //    fNew.IsItalic = iF.IsItalic;
        //    fNew.IsStrikeout = iF.IsStrikeout;
        //    fNew.TypeOffset = iF.TypeOffset;
        //    fNew.Underline = iF.Underline;
        //    ft_AllMain.Add(fNew);
        //    return fNew;
        //}

        //private List<int> cols_AutoWidth;
        //private List<int> cols_CustomWidth;
        //private List<string> cols_Format;
        //private List<string> cols_Type;
        //private List<int> cols_Visible;



        //#region IColor : HSSFColor & XSSFColor
        //private const int iL = 17;
        //private List<short> cl_AllMain_XLS = null;
        //private HSSFPalette cl_Palette_XLS = null;
        //private short cl_Find_XLS(Color c)
        //{
        //    HSSFColor hc = cl_Palette_XLS.FindColor(c.R, c.G, c.B);
        //    if (hc != null && (hc.Indexed <= iL || cl_AllMain_XLS.Contains(hc.Indexed))) return hc.Indexed;
        //    else
        //    {
        //        int ind = iL + cl_AllMain_XLS.Count;
        //        cl_Palette_XLS.SetColorAtIndex(Convert.ToInt16(ind), c.R, c.G, c.B);
        //        cl_AllMain_XLS.Add(Convert.ToInt16(ind));
        //        return cl_AllMain_XLS[cl_AllMain_XLS.Count - 1];
        //    }
        //}
        ////====================================
        //List<XSSFColor> cl_AllMain_XLSX = null;
        //private XSSFColor cl_Find_XLSX(Color cInput)
        //{
        //    foreach (XSSFColor c in cl_AllMain_XLSX)
        //    {
        //        if (c.RGB[0] == cInput.R && c.RGB[1] == cInput.G && c.RGB[2] == cInput.B) return c;
        //    }
        //    XSSFColor cNew = new XSSFColor(cInput);
        //    cl_AllMain_XLSX.Add(cNew);
        //    return cNew;
        //}
        //#endregion
        //private void xl_Init_Main(IWorkbook wb)
        //{
        //    if (wb is HSSFWorkbook)
        //    {
        //        cs_AllMain = new List<ICellStyle>();
        //        ft_AllMain = new List<IFont>();
        //        cl_AllMain_XLS = new List<short>();
        //        cs_Tmp = wb.CreateCellStyle();
        //        ft_Tmp = wb.CreateFont();
        //    }
        //    else
        //    {
        //        cs_AllMain = new List<ICellStyle>();
        //        ft_AllMain_XLSX = new List<XSSFFont>();
        //        cl_AllMain_XLSX = new List<XSSFColor>();
        //        cs_Tmp = wb.CreateCellStyle();
        //        ft_Tmp_XLSX = (XSSFFont)wb.CreateFont();
        //    }
        //}
        ////===============
        //private ICellStyle cs_Tmp = null;
        //private List<ICellStyle> cs_AllMain;
        //private ICellStyle cs_Find(ICellStyle ics, IWorkbook wb)
        //{
        //    foreach (ICellStyle cs in cs_AllMain)
        //    {
        //        if (cs.Index == ics.Index) continue;

        //        if (cs.BorderTop != ics.BorderTop) continue;
        //        if (cs.BorderBottom != ics.BorderBottom) continue;
        //        if (cs.BorderLeft != ics.BorderLeft) continue;
        //        if (cs.BorderRight != ics.BorderRight) continue;

        //        if (cs.FillForegroundColor != ics.FillForegroundColor) continue;
        //        if (cs.FillForegroundColorColor != null && ics.FillForegroundColorColor != null)
        //            if (cs.FillForegroundColorColor.Indexed != ics.FillForegroundColorColor.Indexed) continue;

        //        if (cs.FillBackgroundColor != ics.FillBackgroundColor) continue;
        //        if (cs.FillBackgroundColorColor != null && ics.FillBackgroundColorColor != null)
        //            if (cs.FillBackgroundColorColor.Indexed != ics.FillBackgroundColorColor.Indexed) continue;

        //        if (cs.FontIndex != ics.FontIndex) continue;
        //        if (cs.Alignment != ics.Alignment) continue;
        //        if (cs.VerticalAlignment != ics.VerticalAlignment) continue;
        //        if (cs.WrapText != ics.WrapText) continue;
        //        if (cs.FillPattern != ics.FillPattern) continue;
        //        if (cs.DataFormat != ics.DataFormat) continue;
        //        return cs;
        //    }
        //    ICellStyle csNew = wb.CreateCellStyle();
        //    cs_Copy(ics, csNew);
        //    cs_AllMain.Add(csNew);
        //    return csNew;
        //}
        ////===============
        //#region Export-Excel
        //public void Export(string ExcelFilePath, System.Windows.Forms.ProgressBar proBar)
        //{
        //    System.Globalization.CultureInfo ciOld = System.Windows.Forms.Application.CurrentCulture;
        //    try
        //    {
        //        CultureInfo ciNew = (System.Globalization.CultureInfo)ciOld.Clone();
        //        ciNew.NumberFormat = (NumberFormatInfo)new CultureInfo("en-US", false).NumberFormat.Clone();
        //        System.Windows.Forms.Application.CurrentCulture = ciNew;

        //        Start_Probar(proBar, HeaderArea.Count + dtbMain.Rows.Count + 4 + HandleTotalProgressStep);
        //        if (ExcelFilePath.ToUpper().EndsWith(".XLS"))
        //            ExportExcel_XLS(ExcelFilePath, proBar);
        //        //else if (ExcelFilePath.ToUpper().EndsWith(".XLSX"))
        //        //    ExportExcel_XLSX(ExcelFilePath, proBar);
        //        else
        //            throw new Exception("Unknow File Type!");
        //    }
        //    finally
        //    {
        //        System.Windows.Forms.Application.CurrentCulture = ciOld;
        //        Stop_Probar(proBar);
        //    }
        //}
        //#endregion

        //private void ExportExcel_XLS(string FilePath, System.Windows.Forms.ProgressBar proBar)
        //{ // 1997-2003
        //    #region Init
        //    //// System.Windows.Forms.ProgessBar
        //    //proBar.Visible = false;
        //    //proBar.Style = System.Windows.Forms.ProgessBarStyle.Block;
        //    //proBar.Minimum = 0;
        //    //proBar.Value = proBar.Minimum;
        //    //proBar.Maximum = HeaderArea.Count + (dtbMain != null ? dtbMain.Rows.Count : grd.Rows.Count) + 7 + HandleTo
        //    //proBar.Step = 1;
        //    //proBar.UseWaitCursor = true;
        //    //proBar.Visible = true;
        //    //=========================================================================
        //    int iBRow = HeaderArea.Count;
        //    int iBCol = 0;
        //    int iECol = cols_Visible.Count + iBCol - 1;
        //    int iERow; // ?????
        //    int eRowNow = 0, iCol; // Current Row/Col
        //    //==========================================================================
        //    ////create a entry of DocumentSummaryInformation
        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    cl_AllMain_XLS = new List<short>();
        //    cl_Palette_XLS = wb.GetCustomPalette(); //Init CustomPalette !!!!!!!
        //    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        //    dsi.Company = "BHD";
        //    wb.DocumentSummaryInformation = dsi;
        //    ////create a entry of SummaryInformation
        //    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        //    si.Subject = "BHD-Group";
        //    wb.SummaryInformation = si;
        //    //Insert three sheet just like what Excel does
        //    ISheet s = wb.CreateSheet(System.Windows.Forms.Application.ProductName);
        //    ((HSSFSheet)s).AlternativeFormula = false;
        //    ((HSSFSheet)s).AlternativeExpression = false;
        //    //====================================================================
        //    List<ICellStyle> cs_Cols = new List<ICellStyle>();
        //    IDataFormat fm = wb.CreateDataFormat();
        //    xl_Init_Main(wb);

        //    #region cell style & font Caption
        //    IFont ft_Caption = wb.CreateFont();
        //    ft_Caption.FontName = Default_Font_Name_;
        //    ft_Caption.FontHeightInPoints = Convert.ToInt16(Default_Font_Size_);
        //    ft_Caption.Boldweight = (short)700;
        //    ft_Caption.Color = cl_Find_XLS(Caption_Color_Fore);
        //    //============
        //    ICellStyle cs_Caption = wb.CreateCellStyle();
        //    cs_Caption.Alignment = HorizontalAlignment.Center;
        //    cs_Caption.VerticalAlignment = VerticalAlignment.Center;
        //    cs_Caption.BorderBottom = BorderStyle.Thin;
        //    cs_Caption.BorderLeft = BorderStyle.Thin;
        //    cs_Caption.BorderRight = BorderStyle.Thin;
        //    cs_Caption.BorderTop = BorderStyle.Thin;
        //    cs_Caption.BottomBorderColor = HSSFColor.Black.Index;
        //    cs_Caption.LeftBorderColor = HSSFColor.Black.Index;
        //    cs_Caption.RightBorderColor = HSSFColor.Black.Index;
        //    cs_Caption.TopBorderColor = HSSFColor.Black.Index;
        //    cs_SetBackColor_XLS(cs_Caption, Caption_Color_Back);
        //    cs_Caption.SetFont(ft_Caption);
        //    #endregion

        //    #region cell style & font Default
        //    IFont ft_Default = wb.CreateFont();
        //    ft_Default.FontName = Default_Font_Name_;
        //    ft_Default.FontHeightInPoints = Convert.ToInt16(Default_Font_Size_);
        //    ft_AllMain.Add(ft_Default);
        //    //===
        //    ICellStyle cs_Default = wb.CreateCellStyle();
        //    cs_Default.Alignment = HorizontalAlignment.Left;
        //    cs_Default.VerticalAlignment = VerticalAlignment.Center;
        //    cs_Default.BorderBottom = BorderStyle.Thin;
        //    cs_Default.BorderLeft = BorderStyle.Thin;
        //    cs_Default.BorderRight = BorderStyle.Thin;
        //    cs_Default.BorderTop = BorderStyle.Thin;
        //    cs_Default.BottomBorderColor = HSSFColor.Black.Index;
        //    cs_Default.LeftBorderColor = HSSFColor.Black.Index;
        //    cs_Default.RightBorderColor = HSSFColor.Black.Index;
        //    cs_Default.TopBorderColor = HSSFColor.Black.Index;
        //    cs_Default.SetFont(ft_Default);
        //    cs_Default.DataFormat = 0;
        //    cs_AllMain.Add(cs_Default);
        //    #endregion

        //    #region cell style & font Boolean
        //    IFont ft_Bool = wb.CreateFont();
        //    if (UseCustomBoolean_XL)
        //    {
        //        ft_Bool.FontName = Default_Font_Name_;
        //        ft_Bool.FontHeightInPoints = Convert.ToInt16(Default_Font_Size_);
        //    }
        //    else
        //    {
        //        ft_Bool.FontName = Default_Font_Name_;
        //        ft_Bool.FontHeightInPoints = Convert.ToInt16(Default_Font_Size_);
        //        /*
        //        //ft_Bool.FontName = "Wingdings";
        //        //ft_Bool.FontHeightInPoints = Convert.ToInt16(Default_Font_Size_ + 1);
        //        */
        //    }
        //    ft_AllMain.Add(ft_Bool);
        //    //===
        //    ICellStyle cs_Bool = wb.CreateCellStyle();
        //    cs_Bool.Alignment = HorizontalAlignment.Center;
        //    cs_Bool.VerticalAlignment = VerticalAlignment.Center;
        //    cs_Bool.BorderBottom = BorderStyle.Thin;
        //    cs_Bool.BorderLeft = BorderStyle.Thin;
        //    cs_Bool.BorderRight = BorderStyle.Thin;
        //    cs_Bool.BorderTop = BorderStyle.Thin;
        //    cs_Bool.BottomBorderColor = HSSFColor.Black.Index;
        //    cs_Bool.LeftBorderColor = HSSFColor.Black.Index;
        //    cs_Bool.RightBorderColor = HSSFColor.Black.Index;
        //    cs_Bool.TopBorderColor = HSSFColor.Black.Index;
        //    cs_Bool.SetFont(ft_Bool);
        //    cs_Bool.DataFormat = 0;
        //    cs_AllMain.Add(cs_Bool);
        //    #endregion

        //    //============================================================
        //    IRow r; // Row Tmp
        //    ICell c; // Cell Tmp
        //    string str1; // String Tmp
        //    string str2; // String Tmp
        //    object obj; // tmp
        //    //============================================================
        //    #endregion
        //    #region Header
        //    ipNow++;// proBar.PerformStep();
        //    eRowNow = 0;
        //    for (int i = 0; i < HeaderArea.Count; i++)
        //    {
        //        xlHeader I = HeaderArea[i];
        //        c = s.CreateRow(i).CreateCell(iBCol);
        //        if (I.FontSize != 1923)
        //        {
        //            s.AddMergedRegion(new CellRangeAddress(i, i, iBCol, iECol));
        //            c.SetCellType(CellType.String);
        //            c.SetCellValue(I.Text);
        //            ICellStyle cs = wb.CreateCellStyle();
        //            switch (I.HorizontalAlignment)
        //            {
        //                case HAlign.Left:
        //                    cs.Alignment = HorizontalAlignment.Left;
        //                    break;
        //                case HAlign.Center:
        //                    cs.Alignment = HorizontalAlignment.Center;
        //                    break;
        //                case HAlign.Right:
        //                    cs.Alignment = HorizontalAlignment.Right;
        //                    break;
        //            }
        //            switch (I.VerticalAlignment)
        //            {
        //                case VAlign.Top:
        //                    cs.VerticalAlignment = VerticalAlignment.Top;
        //                    break;
        //                case VAlign.Bottom:
        //                    cs.VerticalAlignment = VerticalAlignment.Bottom;
        //                    break;
        //                case VAlign.Center:
        //                    cs.VerticalAlignment = VerticalAlignment.Center;
        //                    break;
        //            }
        //            IFont ft = wb.CreateFont();
        //            ft.FontName = I.FontName;
        //            ft.FontHeightInPoints = Convert.ToInt16(I.FontSize);
        //            ft.Color = cl_Find_XLS(I.ForeColor);
        //            if (I.FontBold) ft.Boldweight = (short)700;
        //            cs.SetFont(ft);
        //            c.CellStyle = cs;
        //        }
        //        ipNow++;//proBar.PerformStep();
        //        eRowNow++;
        //    }
        //    /*s.CreateRow(iRow).CreareCell(0, CellType.Blank);
        //    s.AddMergedRegion(new CellRangeAddress(iRow, iRow, iBCol, iECol));
        //    iRow++;*/
        //    #endregion
        //    //============================================================
        //    #region Data Table
        //    if (dtbMain != null)
        //    {
        //        r = s.CreateRow(eRowNow);
        //        iCol = iBCol;
        //        for (int j = 0; j < dtbMain.Columns.Count; j++)
        //        {
        //            if (cols_Visible.Contains(j) == false) continue;
        //            c = r.CreateCell(iCol, CellType.String);
        //            c.CellStyle = cs_Caption;
        //            c.SetCellValue(dtbMain.Columns[j].Caption);
        //            Set_ColumnAutoWidth(j, dtbMain.Columns[j].Caption, true);
        //            iCol++;
        //        }
        //        ipNow++;//proBar.PerformStep();
        //        eRowNow++;
        //        cs_Cols = new List<ICellStyle>();
        //        for (int j = 0; j < dtbMain.Columns.Count; j++)
        //        {
        //            if (cols_Visible.Contains(j) == false) { cs_Cols.Add(null); continue; }
        //            ICellStyle cs = wb.CreateCellStyle();
        //            cs.Alignment = HorizontalAlignment.General;
        //            cs.VerticalAlignment = VerticalAlignment.Center;
        //            cs.BorderBottom = BorderStyle.Thin;
        //            cs.BorderLeft = BorderStyle.Thin;
        //            cs.BorderRight = BorderStyle.Thin;
        //            cs.BorderTop = BorderStyle.Thin;
        //            cs.BottomBorderColor = HSSFColor.Black.Index;
        //            cs.LeftBorderColor = HSSFColor.Black.Index;
        //            cs.RightBorderColor = HSSFColor.Black.Index;
        //            cs.TopBorderColor = HSSFColor.Black.Index;
        //            /*cs.ShrinkToFit = true;*/
        //            cs.SetFont(ft_Default);
        //            if (cols_Format[j] != "") cs.DataFormat = fm.GetFormat(cols_Format[j]);
        //            else cs.DataFormat = 0;
        //            cs_Cols.Add(cs);
        //        }

        //        for (int i = 0; i < dtbMain.Rows.Count; i++)
        //        {
        //            if (eRowNow > 65535)
        //            {
        //                s = wb.CreateSheet((wb.GetSheetName(0) + "_" + (wb.GetSheetIndex(s.SheetName) + 2).ToString()));
        //                eRowNow = 0;
        //            }
        //            r = s.CreateRow(eRowNow);
        //            iCol = iBCol;
        //            for (int j = 0; j < dtbMain.Columns.Count; j++)
        //            {
        //                if (!cols_Visible.Contains(j)) continue;
        //                str1 = cols_Type[j];
        //                obj = dtbMain.Rows[i][j];
        //                if (obj != DBNull.Value)
        //                {
        //                    if (str1 == tString || str1 == tDateTime)
        //                    {
        //                        c = r.CreateCell(iCol, CellType.String);
        //                        c.SetCellValue(obj.ToString());
        //                    }
        //                    else if (str1 == tInt16 || str1 == tInt32 || str1 == tInt64 || str1 == tFloat || str1 == tDouble || str1 == tDecimal)
        //                    {
        //                        c = r.CreateCell(iCol, CellType.Numeric);
        //                        c.SetCellValue(Convert.ToDouble(obj));
        //                    }
        //                    else if (str1 == tBool)
        //                    {
        //                        c = r.CreateCell(iCol, CellType.Boolean);
        //                        //c.SetCellValue(Convert.ToBoolean(obj));
        //                        c.SetCellValue(Convert.ToBoolean(obj) ? _CheckSymbol : _UnCheckSymbol);
        //                    }
        //                    /*else if (str == tDataTime)
        //                    {
        //                        c = r.CreateCell(iCol);
        //                        c.SetCellValue(Convert.ToDateTime(obj));
        //                    }*/
        //                    else
        //                    {
        //                        c = r.CreateCell(iCol);
        //                        c.SetCellValue(obj.ToString());
        //                    }
        //                    Set_ColumnAutoWidth(j, obj.ToString(), false);
        //                }
        //                else c = r.CreateCell(iCol, CellType.Blank);
        //                c.CellStyle = cs_Cols[j];
        //                iCol++;
        //            }
        //            ipNow++;//proBar.PerformStep();
        //            eRowNow++;
        //        }
        //    }
        //    #endregion
            
        //    iERow = eRowNow - 1;
        //    #region Set Column Width
        //    int FixWidth = 36;
        //    for (int i = 0; i < wb.NumberOfSheets; i++)
        //    {
        //        s = wb.GetSheetAt(i);
        //        iCol = iBCol;
        //        for (int j = 0; j < cols_CustomWidth.Count; j++)
        //        {
        //            try
        //            {
        //                if (!cols_Visible.Contains(j)) continue;
        //                if (cols_CustomWidth[j] != 0)
        //                    s.SetColumnWidth(iCol, cols_CustomWidth[j] * FixWidth);
        //                else
        //                    s.SetColumnWidth(iCol, cols_AutoWidth[j] * FixWidth);
        //            }
        //            catch { s.AutoSizeColumn(iCol); }
        //            iCol++;
        //        }
        //    }
        //    #endregion
        //    ipNow++;//proBar.PerformStep();

        //    #region Handle-Region
        //    if (CustomExcel != null)
        //    {
        //        IsHandling = true;
        //        HandleWorkbook = wb;
        //        HandleSheet = s;
        //        HandleICellStypeList = cs_Cols;
        //        CustomExcel(this, new EventArgs());
        //        IsHandling = false;
        //    }
        //    ipNow++;//proBar.PerformStep();
        //    #endregion

        //    #region FOR: Fomula
        //    foreach (Formular_Struct fo in Formular_Arr)
        //    {
        //        s = wb.GetSheetAt(fo.SheetIndex);
        //        r = s.GetRow(fo.RowIndex);
        //        c = r.Cells[fo.ColIndex];
        //        c.SetCellFormula(fo.ExcelFormular);
        //        //wb.GetSheetAt(fo.SheetIndex).GetRow(fo.RowIndex).Cell[fo.ColIndex]
        //    }
        //    #endregion

        //    wb.SetActiveSheet(0);
        //    //Write the stream data of workbook
        //    //if (File.Exists(FilePath)) File.Delete(FilePath);
        //    FileStream fs = new FileInfo(FilePath).Open(FileMode.Create, FileAccess.ReadWrite);
        //    wb.Write(fs);
        //    ipNow++;//proBar.PerFormStep();
        //    fs.Close();

        //    //proBar.Value = proBar.Maximum - 1;
        //    //proBar.PerformStep();
        //    //proBar.Visible = false;
        //    //proBar.UseWaitCursor = false;
        //}

        //#region For: Handle
        //public object HandleICellStypeList;
        //public int HandleTotalProgressStep = 0;
        //public object HandleWorkbook;
        //public object HandleSheet;
        //public event EventHandler CustomExcel;
        //private bool IsHandling = false;
        //#endregion

        //#region FOR: ProgressBar
        //private System.Windows.Forms.ProgressBar proBar;

        //private int ipMaxV = 0; //Max Value
        //private int ipMaxP = 200; //Max Percent //Fixed
        //private int _ipNow = 0; //Current Value
        //private int ipNow
        //{
        //    get { return _ipNow; }
        //    set
        //    {
        //        _ipNow = value;
        //        if ((_ipNow * ipMaxP / ipMaxV) >= proBar.Value + 1)
        //        {
        //            proBar.PerformStep();
        //            System.Windows.Forms.Application.DoEvents();
        //        }
        //    }
        //}

        //private void Start_Probar(System.Windows.Forms.ProgressBar progressBar, int MaxPlus)
        //{
        //    if (MaxPlus < ipMaxP) ipMaxP = MaxPlus;
        //    if (progressBar.Value >= progressBar.Maximum - 1)
        //        progressBar.Value = progressBar.Minimum;
        //    MaxPlus++;
        //    int iPercent = 0, iNewValue = 0;
        //    if (progressBar.Maximum != progressBar.Minimum)
        //    {
        //        iPercent = Convert.ToInt32(((progressBar.Value - progressBar.Minimum) * ipMaxP) /
        //            (progressBar.Maximum - progressBar.Minimum));
        //        iNewValue = Convert.ToInt32((MaxPlus / (ipMaxP - iPercent)) * iPercent);
        //    }
        //    progressBar.Style = System.Windows.Forms.ProgressBarStyle.Blocks;
        //    progressBar.Minimum = 0;
        //    ipMaxV = MaxPlus + iNewValue;
        //    //progressBar.Value = progressBar.Minimum;
        //    progressBar.Maximum = ipMaxP;
        //    progressBar.Value = iPercent;
        //    progressBar.Step = 1;
        //    progressBar.PerformStep();

        //    progressBar.Show();
        //    progressBar.BringToFront();
        //    proBar = progressBar;
        //    progressBar.Cursor = System.Windows.Forms.Cursors.WaitCursor;
        //    System.Windows.Forms.Form frm = progressBar.FindForm();
        //    if (frm != null)
        //    {
        //        frm.UseWaitCursor = true;
        //    }
        //    progressBar.PerformStep();
        //}

        //private static void Stop_Probar(System.Windows.Forms.ProgressBar progressBar)
        //{
        //    progressBar.PerformStep();
        //    progressBar.Hide();
        //    progressBar.Cursor = System.Windows.Forms.Cursors.Default;
        //    System.Windows.Forms.Form frm = progressBar.FindForm();
        //    if (frm != null)
        //    {
        //        frm.UseWaitCursor = false;
        //    }
        //    progressBar.Value = progressBar.Maximum;
        //    progressBar.PerformStep();
        //}

        //public void OpenExcel(string ExcelFilePath)
        //{
        //    System.Diagnostics.Process.Start(ExcelFilePath);
        //}
        //#endregion


        //#region Cách sử dụng
        ////private void sb_ExportExcel() {
        ////    string strPathExcel = "";
        ////    System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
        ////    sfd.FileName = string.Format("{0}_{1:yyyyMMddhhmm}.xls", this.Name, DateTime.Now);
        ////    sfd.Filter = "Excel File (*.xls)|*.xls|All File(*.*)|*.*";
        ////    if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        ////        strPathExcel = sfd.FileName;
        ////    if (strPathExcel.Trim() == "")
        ////        return;
        ////    C1FlexGrid grdExcel = grd;
        ////    bhdExcel clsExcel = new bhdExcel(grdExcel);
        ////    clsExcel.UseCellBackColor = true;
        ////    clsExcel.UseCellForeColor = true;
        ////    clsExcel.UseCellFontStyle = true;
        //////    clsExcel.Set_Header("", 0);
        ////    clsExcel.Set_Header(this.Text, 1);
        //////    clsExcel.Set_Header("", 0);
        ////    clsExcel.HideColumn((int)DataCol.No);
        ////    clsExcel.Export(strPathExcel, progressBar);
        ////    clsExcel.OpenExcel(strPathExcel);
        ////}
        //#endregion
    }
}
