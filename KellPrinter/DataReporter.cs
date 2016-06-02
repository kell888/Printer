using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;

namespace KellPrinter
{
    #region 定义单元格常用到样式的枚举
    public enum XlsStyle
    {
        Title,
        Header,
        Bottom,
        Serial,
        Url,
        Time,
        Number,
        Money,
        Percent,
        Chupper,
        Tnumber,
        Default
    }
    #endregion
    public class DataReporter
    {
        #region
        //image size
        int _Width = 600;
        int _Height = 420;
        //pager
        private int _TopMargin = 50;
        private int _LeftMargin = 60;
        private int _RightMargin = 50;
        private int _BottomMargin = 60;
        private Font _TitleFont = new Font("宋体", 15, FontStyle.Bold);
        private Font _ColumnsHeaderFont = new Font("宋体", 10, FontStyle.Bold);
        private Font _ContentFont = new Font("宋体", 9, FontStyle.Regular);
        private Font _BottomFont = new Font("宋体", 10, FontStyle.Bold);
        private SolidBrush brush = new SolidBrush(Color.Black);
        private Pen pen = new Pen(new SolidBrush(Color.Black));
        private int _RowHeight = 30;
        private int _CurrentPageIndex;
        private int _PageCount;
        private int _RowsCount;
        private int _CurrentRowsIndex;
        private int _MaxRowsCount = 35;
        /// <summary>
        /// 获取或设置每页的最大行数，默认为35行
        /// </summary>
        public int MaxRowsCountPerPage
        {
            get { return _MaxRowsCount; }
            set { _MaxRowsCount = value; }
        }
        private Point _CurrentPoint;
        private DataTable _DT;
        private string _Title;
        private string _ImgTitle;
        private string[] _ColumnsHeader;
        private string[] _BottomStr;

        public string[] BottomStr
        {
            get { return _BottomStr; }
            set { _BottomStr = value; }
        }
        #endregion

        #region
        public DataReporter(DataTable data, PrintArgs args = null)
        {
            if (data == null)
                throw new Exception("要打印导出的数据表不能回为空！请确认参数data=null.");

            if (args != null)
            {
                _Title = args.Title;
                _ImgTitle = args.ImgTitle;
                _ColumnsHeader = args.ColumnsHeader;
                _BottomStr = args.BottomStr;
            }
            else
            {
                _Title = data.TableName;
                _ImgTitle = data.TableName + "饼状图";
                List<string> cols = new List<string>();
                cols.Add("序号");
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    cols.Add(data.Columns[i].ColumnName);
                }
                _ColumnsHeader = cols.ToArray();
            }
            _DT = data;
            _RowsCount = data.Rows.Count;
            _CurrentPageIndex = 0;
            _CurrentRowsIndex = 0;
            //pagecount
            if ((data.Rows.Count + 20) % _MaxRowsCount == 0)
                _PageCount = (data.Rows.Count + 20) / _MaxRowsCount;
            else
                _PageCount = ((data.Rows.Count + 20) / _MaxRowsCount) + 1;
        }
        #endregion

        #region
        public Exception SaveAsExcel(string fullname = null, bool showBottomUnderline = true)
        {
            Exception e = null;
            if (_ColumnsHeader.Length < _DT.Columns.Count || _ColumnsHeader.Length > _DT.Columns.Count + 1)
                return new Exception("列头数目与数据不符！注意：HeaderCount=[DataColumnCount, DataColumnCount + 1]");
            string reportName = "报表";
            if (!string.IsNullOrEmpty(_DT.TableName))
                reportName = _DT.TableName;
            if (string.IsNullOrEmpty(fullname))
            {
                string exportPath = Directory.GetCurrentDirectory() + "\\Reports\\";
                if (!Directory.Exists(exportPath))
                    Directory.CreateDirectory(exportPath);
                string filename = reportName + "_" + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss");
                fullname = exportPath + filename + ".xls";
            }

            IWorkbook wb = null;
            try
            {
                if (fullname.EndsWith(".xlsx", StringComparison.InvariantCultureIgnoreCase)) // 2007版本及其以上的版本
                    wb = new XSSFWorkbook();
                else if (fullname.EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase)) // 2003版本
                    wb = new HSSFWorkbook();
                if (wb == null)
                    return new Exception("无法创建Workbook！");

                //创建表  
                ISheet sh = wb.CreateSheet(reportName);
                //创建标题
                bool showTile = !string.IsNullOrEmpty(_Title);
                if (showTile)
                {
                    ICell ce;
                    int colCount = _ColumnsHeader.Length;
                    IRow row = sh.CreateRow(0);
                    row.Height = (short)(20 * _TitleFont.Height);
                    for (int i = 0; i < colCount; i++)
                    {
                        ce = row.CreateCell(i, CellType.BLANK);
                        if (i == 0)
                        {
                            ce.SetCellValue(_Title);
                            ce.CellStyle = GetCellStyle(wb, XlsStyle.Title, null, null, _TitleFont);
                        }
                    }
                    sh.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, colCount - 1));//该方法的参数次序是：开始行号，结束行号，开始列号，结束列号。
                }
                ICell icell;
                //创建表头
                bool hasSerailCol = _ColumnsHeader.Length == _DT.Columns.Count + 1;
                int headRowNum = 0;
                if (showTile) headRowNum = 1;
                IRow row0 = sh.CreateRow(headRowNum);
                row0.Height = (short)(20 * _ColumnsHeaderFont.Height);//20 * 20;
                //设置表头
                for (int i = 0; i < _ColumnsHeader.Length; i++)
                {
                    if (hasSerailCol && i == 0)
                    {
                        sh.SetColumnWidth(i, 10 * 256);
                    }
                    else
                    {
                        sh.SetColumnWidth(i, 15 * 256);
                    }
                    icell = row0.CreateCell(i);
                    icell.CellStyle = GetCellStyle(wb, XlsStyle.Header, _ColumnsHeaderFont);
                    icell.SetCellValue(_ColumnsHeader[i]);
                }
                if (_DT.Rows.Count > 0)
                {
                    int begin=1;
                    if (showTile) begin = 2;
                    ICell cell;
                    int num = 0;
                    for (int j = 0; j < _DT.Rows.Count; j++)
                    {
                        num++;
                        //创建数据行
                        DataRow dr = _DT.Rows[j];
                        IRow row1 = sh.CreateRow(j + begin);
                        row1.Height = (short)(20 * _ContentFont.Height);//20 * 15;
                        for (int k = 0; k < _ColumnsHeader.Length; k++)
                        {
                            CellType ct = CellType.BLANK;
                            string val = "";
                            decimal d;
                            bool b;
                            DateTime t;
                            if (k > 0)
                            {
                                val = dr[k - 1].ToString();
                                if (bool.TryParse(val, out b))
                                    ct = CellType.BOOLEAN;
                                else if (DateTime.TryParse(val, out t))
                                    ct = CellType.STRING;
                                else if (decimal.TryParse(val, out d))
                                    ct = CellType.NUMERIC;
                            }
                            cell = row1.CreateCell(k, ct);
                            if (k == 0)
                            {
                                cell.CellStyle = GetCellStyle(wb, XlsStyle.Serial, null, _ContentFont);
                                cell.SetCellValue(num);
                                continue;
                            }
                            if (DateTime.TryParse(val, out t))
                                cell.CellStyle = GetCellStyle(wb, XlsStyle.Time, null, _ContentFont);
                            else if (decimal.TryParse(val, out d))
                                cell.CellStyle = GetCellStyle(wb, XlsStyle.Number, null, _ContentFont);
                            else
                                cell.CellStyle = GetCellStyle(wb, XlsStyle.Default, null, _ContentFont);
                            cell.SetCellValue(val);
                        }
                    }
                    //设置表尾
                    if (_BottomStr != null)
                    {
                        int tail = begin + num;
                        IRow row2 = sh.CreateRow(tail);
                        row2.Height = (short)(20 * _BottomFont.Height);//20 * 20;
                        ICell cel;
                        for (int i = 0; i < _BottomStr.Length; i++)
                        {
                            string bot = _BottomStr[i];
                            cel = row2.CreateCell(i * 2, CellType.STRING);
                            cel.CellStyle = GetCellStyle(wb, XlsStyle.Bottom, null, null, null, _BottomFont);
                            cel.SetCellValue(bot);
                            if (showBottomUnderline)
                            {
                                cel = row2.CreateCell(i * 2 + 1, CellType.BLANK);
                                ICellStyle cellStyle = wb.CreateCellStyle();
                                cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                                cel.CellStyle = cellStyle;
                            }
                        }
                    }
                }
                using (FileStream stm = File.OpenWrite(fullname))
                {
                    wb.Write(stm);
                }
            }
            catch (Exception ex)
            {
                e = ex;
            }
            finally
            {
                wb = null;
            }

            return e;
        }
        #endregion        

        #region 定义单元格常用到样式
        private static ICellStyle GetCellStyle(IWorkbook wb, XlsStyle str, Font headFont = null, Font contentFont = null, Font titleFont = null, Font bottomFont = null)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            //定义几种字体  
            //也可以一种字体，写一些公共属性，然后在下面需要时加特殊的

            IFont font = wb.CreateFont();
            font.FontName = "微软雅黑";
            font.FontHeightInPoints = 10;

            if (contentFont != null)
            {
                font.FontName = contentFont.Name;
                if (contentFont.Bold) font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD;
                font.FontHeightInPoints = (short)contentFont.SizeInPoints;
                font.IsItalic = contentFont.Italic;
                font.IsStrikeout = contentFont.Strikeout;
                font.Underline = (byte)(contentFont.Underline ? 1 : 0);
            }

            IFont linkAddresFont = wb.CreateFont();
            linkAddresFont.Color = HSSFColor.OLIVE_GREEN.BLUE.index;
            linkAddresFont.IsItalic = true;//下划线  
            linkAddresFont.FontName = "微软雅黑";

            //边框  
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
            //边框颜色  
            //cellStyle.BottomBorderColor = HSSFColor.OLIVE_GREEN.BLUE.index;
            //cellStyle.TopBorderColor = HSSFColor.OLIVE_GREEN.BLUE.index;
            //背景图形，我没有用到过。感觉很丑  
            //cellStyle.FillBackgroundColor = HSSFColor.OLIVE_GREEN.BLUE.index;  
            //cellStyle.FillForegroundColor = HSSFColor.OLIVE_GREEN.BLUE.index;  
            cellStyle.FillForegroundColor = HSSFColor.WHITE.index;
            // cellStyle.FillPattern = FillPatternType.NO_FILL;  
            cellStyle.FillBackgroundColor = HSSFColor.BLUE.index;
            //水平对齐  
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.LEFT;
            //垂直对齐  
            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            //自动换行  
            cellStyle.WrapText = true;
            //缩进;当设置为1时，前面留的空白太大了。希旺官网改进。或者是我设置的不对  
            cellStyle.Indention = 0;
            //上面基本都是设共公的设置  
            //下面列出了常用的字段类型  
            switch (str)
            {
                case XlsStyle.Title:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                    if (titleFont != null)
                    {
                        font.FontName = titleFont.Name;
                        if (titleFont.Bold) font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD;
                        font.FontHeightInPoints = (short)titleFont.SizeInPoints;
                        font.IsItalic = titleFont.Italic;
                        font.IsStrikeout = titleFont.Strikeout;
                        font.Underline = (byte)(titleFont.Underline ? 1 : 0);
                    }
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Header:
                    // cellStyle.FillPattern = FillPatternType.LEAST_DOTS;
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                    if (headFont != null)
                    {
                        font.FontName = headFont.Name;
                        if (headFont.Bold) font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD;
                        font.FontHeightInPoints = (short)headFont.SizeInPoints;
                        font.IsItalic = headFont.Italic;
                        font.IsStrikeout = headFont.Strikeout;
                        font.Underline = (byte)(headFont.Underline ? 1 : 0);
                    }
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.MEDIUM;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.MEDIUM;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.MEDIUM;
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.MEDIUM;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Bottom:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.RIGHT;
                    if (bottomFont != null)
                    {
                        font.FontName = bottomFont.Name;
                        if (bottomFont.Bold) font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD;
                        font.FontHeightInPoints = (short)bottomFont.SizeInPoints;
                        font.IsItalic = bottomFont.Italic;
                        font.IsStrikeout = bottomFont.Strikeout;
                        font.Underline = (byte)(bottomFont.Underline ? 1 : 0);
                    }
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.NONE;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Serial:
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                    font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Time:
                    IDataFormat dataStyle = wb.CreateDataFormat();
                    cellStyle.DataFormat = dataStyle.GetFormat("yyyy-mm-dd");
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Number:
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.GENERAL;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Money:
                    IDataFormat format = wb.CreateDataFormat();
                    cellStyle.DataFormat = format.GetFormat("￥#,##0");
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.GENERAL;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Url:
                    linkAddresFont.Underline = 1;
                    cellStyle.SetFont(linkAddresFont);
                    break;
                case XlsStyle.Percent:
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
                    cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.GENERAL;
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Chupper:
                    IDataFormat format1 = wb.CreateDataFormat();
                    cellStyle.DataFormat = format1.GetFormat("[DbNum2][$-804]0");
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Tnumber:
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00E+00");
                    cellStyle.SetFont(font);
                    break;
                case XlsStyle.Default:
                    cellStyle.SetFont(font);
                    break;
            }
            return cellStyle;

        }
        #endregion

        #region//对dt排序
        public static DataTable Sort(DataTable dataTable, Dictionary<int,SortOrder> colSorts)
        {
            bool hasSort = false;
            List<string> orderNames = new List<string>();
            foreach (KeyValuePair<int, SortOrder> sort in colSorts)
            {
                if (sort.Value == SortOrder.None) continue;
                hasSort = true;
                string s = " ASC";
                if (sort.Value == SortOrder.Descending)
                    s = " DESC";
                string orderName = dataTable.Columns[sort.Key].ColumnName + s;
                orderNames.Add(orderName);
            }
            if (hasSort)
            {
                DataView dv = dataTable.DefaultView;
                dv.Sort = string.Join(",", orderNames.ToArray());
                DataTable dt = dv.ToTable();
                return dt;
            }
            return dataTable;
        }
        #endregion

        #region 打印报表
        public void PrintReport()
        {
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrintPage += PrintPage;
            printDoc.BeginPrint += BeginPrint;
            PrintPreviewDialog pPreviewDialog = new PrintPreviewDialog();
            pPreviewDialog.Document = printDoc;
            pPreviewDialog.ShowIcon = false;
            pPreviewDialog.PrintPreviewControl.Zoom = 1.0;
            pPreviewDialog.TopLevel = false;
            SetPrintPreviewDialog(pPreviewDialog);
            PrintDialog pd = new PrintDialog();
            pd.Document = pPreviewDialog.Document;
            pd.UseEXDialog = true;
            if (pd.ShowDialog() == DialogResult.OK)
                pPreviewDialog.Document.Print();
        }
        #endregion

        #region 预览打印报表
        public PrintPreviewDialog PreviewPrintReport()
        {
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrintPage += PrintPage;
            printDoc.BeginPrint += BeginPrint;
            PrintPreviewDialog pPreviewDialog = new PrintPreviewDialog();
            pPreviewDialog.Document = printDoc;
            pPreviewDialog.ShowIcon = false;
            pPreviewDialog.PrintPreviewControl.Zoom = 1.0;
            pPreviewDialog.TopLevel = false;
            SetPrintPreviewDialog(pPreviewDialog);
            return pPreviewDialog;
        }
        #endregion

        #region

        #region 绘制饼状图
        ///<summary>
        /// 绘制饼状图
        ///</summary>
        ///<returns></returns>
        private Bitmap GetPieImage(string title, DataTable dataTable)
        {
            Bitmap image = GenerateImage(title);
            dataTable = DataFormat(dataTable);
            //主区域图形
            Rectangle RMain = new Rectangle(35, 70, 380, 300);
            //图例信息
            Rectangle RDes = new Rectangle(445, 90, 10, 10);
            Font f = new Font("宋体", 10, FontStyle.Regular);

            Graphics g = Graphics.FromImage(image);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            try
            {
                //分析数据，绘制饼图和图例说明
                double[] ItemRate = GetItemRate(dataTable);
                int[] ItemAngle = GetItemAngle(ItemRate);
                int Angle1 = 0;
                int Angle2 = 0;
                int len = ItemRate.Length;
                Color c = new Color();
                //3D
                g.DrawPie(new Pen(Color.Black), RMain, 0F, 360F);
                g.DrawPie(new Pen(Color.Black), new Rectangle(RMain.X, RMain.Y + 10, RMain.Width, RMain.Height), 0F, 360F);
                g.FillPie(new SolidBrush(Color.Black), new Rectangle(RMain.X, RMain.Y + 10, RMain.Width, RMain.Height), 0F, 360F);
                //绘制
                for (int i = 0; i < len; i++)
                {
                    Angle2 = ItemAngle[i];
                    //if (c != GetRandomColor(i))
                    c = GetRandomColor(i);

                    SolidBrush brush = new SolidBrush(c);
                    string DesStr = dataTable.Rows[i][0].ToString() + "(" + (ItemRate[i] * 100).ToString(".00") + "%" + ")";
                    //
                    DrawPie(image, RMain, c, Angle1, Angle2);
                    Angle1 += Angle2;
                    DrawDes(image, RDes, c, DesStr, f, i);
                }

                return image;
            }
            finally
            {
                g.Dispose();
            }
        }
        #endregion

        #region 绘制图像的基本数据计算方法
        ///<summary>
        /// 数据格式化
        ///</summary>
        private DataTable DataFormat(DataTable dataTable)
        {
            if (dataTable == null)
                return dataTable;
            //把大于等于10的行合并，
            if (dataTable.Rows.Count <= 10)
                return dataTable;
            //new Table
            DataTable dataTableNew = dataTable.Copy();
            dataTableNew.Rows.Clear();
            for (int i = 0; i < 8; i++)
            {
                DataRow dataRow = dataTableNew.NewRow();
                dataRow[0] = dataTable.Rows[i][0];
                dataRow[1] = dataTable.Rows[i][1];
                dataTableNew.Rows.Add(dataRow);
            }
            DataRow dr = dataTableNew.NewRow();
            dr[0] = "其它";
            double allValue = 0;
            for (int i = 9; i < dataTable.Rows.Count; i++)
            {
                allValue += Convert.ToDouble(dataTable.Rows[i][1]);
            }
            dr[1] = allValue;
            dataTableNew.Rows.Add(dr);
            return dataTableNew;
        }
        ///<summary>
        /// 计算数值总和
        ///</summary>
        private static double Sum(DataTable dataTable)
        {
            double t = 0;
            foreach (DataRow dr in dataTable.Rows)
            {
                t += Convert.ToDouble(dr[1]);
            }
            return t;
        }
        ///<summary>
        /// 计算各项比例
        ///</summary>
        private static double[] GetItemRate(DataTable dataTable)
        {
            double sum = Sum(dataTable);
            double[] ItemRate = new double[dataTable.Rows.Count];
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                ItemRate[i] = Convert.ToDouble(dataTable.Rows[i][1]) / sum;
            }
            return ItemRate;
        }
        ///<summary>
        /// 根据比例，计算各项角度值
        ///</summary>
        private static int[] GetItemAngle(double[] ItemRate)
        {
            int[] ItemAngel = new int[ItemRate.Length];
            for (int i = 0; i < ItemRate.Length; i++)
            {
                double t = 360 * ItemRate[i];
                ItemAngel[i] = Convert.ToInt32(t);
            }
            return ItemAngel;
        }
        #endregion

        #region// 随即扇形区域颜色，绘制区域框，
        ///<summary>
        /// 生成随机颜色
        ///</summary>
        ///<returns></returns>
        private static Color GetRandomColor(int seed)
        {
            Random random = new Random(seed);
            int r = 0;
            int g = 0;
            int b = 0;
            r = random.Next(0, 230);
            g = random.Next(0, 230);
            b = random.Next(0, 235);
            Color randomcolor = Color.FromArgb(r, g, b);
            return randomcolor;
        }
        ///<summary>
        /// 绘制区域框、阴影
        ///</summary>
        private static Bitmap DrawRectangle(Bitmap image, Rectangle rect)
        {
            Bitmap Image = image;
            Graphics g = Graphics.FromImage(Image);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            try
            {
                Rectangle rn = new Rectangle(rect.X + 3, rect.Y + 3, rect.Width, rect.Height);
                SolidBrush brush1 = new SolidBrush(Color.FromArgb(233, 234, 249));
                SolidBrush brush2 = new SolidBrush(Color.FromArgb(221, 213, 215));
                //
                g.FillRectangle(brush2, rn);
                g.FillRectangle(brush1, rect);
                return Image;
            }
            finally
            {
                g.Dispose();
            }
        }
        #endregion

        #region 绘制图例框、图列信息
        ///<summary>
        /// 绘制图例信息
        ///</summary>
        private static Bitmap DrawDes(Bitmap image, Rectangle rect, Color c, string DesStr, Font f, int i)
        {
            Bitmap Image = image;
            Graphics g = Graphics.FromImage(Image);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
            try
            {
                SolidBrush brush = new SolidBrush(c);
                Rectangle R = new Rectangle(rect.X, rect.Y + 25 * i, rect.Width, rect.Height);
                Point p = new Point(rect.X + 12, rect.Y + 25 * i);
                //❀颜色矩形框
                g.FillRectangle(brush, R);
                //文字说明
                g.DrawString(DesStr, f, new SolidBrush(Color.Black), p);
                return Image;
            }
            finally
            {
                g.Dispose();
            }
        }
        //绘制扇形
        private static Bitmap DrawPie(Bitmap image, Rectangle rect, Color c, int Angle1, int Angle2)
        {
            Bitmap Image = image;
            Graphics g = Graphics.FromImage(Image);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            try
            {
                SolidBrush brush = new SolidBrush(c);
                Rectangle R = new Rectangle(rect.X, rect.Y, rect.Width, rect.Height);
                g.FillPie(brush, R, Angle1, Angle2);
                return Image;
            }
            finally
            {
                g.Dispose();
            }
        }
        #endregion

        #region 绘制基本图形
        ///<summary>
        /// 生成图片，统一设置图片大小、背景色,图片布局，及标题
        ///</summary>
        ///<returns>图片</returns>
        private Bitmap GenerateImage(string Title)
        {
            Bitmap image = new Bitmap(_Width, _Height);
            Graphics g = Graphics.FromImage(image);
            //标题
            Point PTitle = new Point(30, 20);
            Font f1 = new Font("黑体", 12, FontStyle.Bold);
            //线
            int len = (int)g.MeasureString(Title, f1).Width;
            Point PLine1 = new Point(20, 40);
            Point PLine2 = new Point(20 + len + 20, 40);
            Pen pen = new Pen(new SolidBrush(Color.FromArgb(8, 34, 231)), 1.5f);
            //主区域,主区域图形
            Rectangle RMain1 = new Rectangle(20, 55, 410, 345);
            Rectangle RMain2 = new Rectangle(25, 60, 400, 335);
            //图例区域
            Rectangle RDes1 = new Rectangle(440, 55, 150, 345);
            //图例说明
            string Des = "图例说明：";
            Font f2 = new Font("黑体", 10, FontStyle.Bold);
            Point PDes = new Point(442, 65);
            //图例信息，后面的x坐标上累加20
            Rectangle RDes2 = new Rectangle(445, 90, 10, 10);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            try
            {
                //设置背景色、绘制边框
                g.Clear(Color.White);
                g.DrawRectangle(pen, 1, 1, _Width - 2, _Height - 2);
                //绘制标题、线
                if (!string.IsNullOrEmpty(Title))
                {
                    g.DrawString(Title, f1, new SolidBrush(Color.Black), PTitle);
                    g.DrawLine(pen, PLine1, PLine2);
                }
                //主区域 
                image = DrawRectangle(image, RMain1);
                //图例区域
                image = DrawRectangle(image, RDes1);
                //“图例说明”
                g.DrawString(Des, f2, new SolidBrush(Color.Black), PDes);
                //return 
                return image;
            }
            finally
            {
                g.Dispose();
            }

        }
        #endregion

        #endregion

        #region 绘制图形、报表

        #region print Event
        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            _CurrentPageIndex++;
            _CurrentPoint = new Point(_LeftMargin, _RightMargin);
            int serialNumWidth = 60;
            int colWidth = (e.PageBounds.Width - _LeftMargin - _RightMargin - serialNumWidth) / _DT.Columns.Count;
            //第一页绘制标题，图形
            if (_CurrentPageIndex == 1)
            {
                DrawTitle(e);
                if (_DT.Columns.Count > 1 && _DT.Rows.Count > 0)
                {
                    object o = _DT.Rows[0][1];
                    double d;
                    if (double.TryParse(o.ToString(), out d))
                        DrawImage(e);
                }
                DrawTableHeader(e, serialNumWidth, colWidth);
                DrawBottom(e);
                DrawTableAndSerialNumAndData(e, serialNumWidth, colWidth);
                if (_PageCount > 1)
                    e.HasMorePages = true;

            }
            else if (_CurrentPageIndex == _PageCount)
            {
                DrawTableHeader(e, serialNumWidth, colWidth);
                DrawTableAndSerialNumAndData(e, serialNumWidth, colWidth);
                DrawBottom(e);
                e.HasMorePages = false;
                e.Cancel = true;
            }
            else
            {
                DrawTableHeader(e, serialNumWidth, colWidth);
                DrawTableAndSerialNumAndData(e, serialNumWidth, colWidth);
                DrawBottom(e);
                e.HasMorePages = true;

            }
        }
        private void BeginPrint(object sender, PrintEventArgs e)
        {
            _CurrentPageIndex = 0;
            _CurrentRowsIndex = 0;
            e.Cancel = false;
        }
        #endregion

        #region 绘制标题
        private void DrawTitle(PrintPageEventArgs e)
        {
            //标题 居中
            if (!string.IsNullOrEmpty(_Title))
            {
                _CurrentPoint.X = (e.PageBounds.Width) / 2 - (int)(e.Graphics.MeasureString(_Title, _TitleFont).Width) / 2;
                e.Graphics.DrawString(_Title, _TitleFont, new SolidBrush(Color.Black), _CurrentPoint);
                _CurrentPoint.Y += (int)(e.Graphics.MeasureString(_Title, _TitleFont).Height);
                //标题下的线
                int len = (int)(e.Graphics.MeasureString(_Title, _TitleFont).Width) + 100;
                int start = (e.PageBounds.Width) / 2 - len / 2;
                e.Graphics.DrawLine(new Pen(new SolidBrush(Color.Black)), new Point(start, _CurrentPoint.Y), new Point(start + len, _CurrentPoint.Y));
                _CurrentPoint.Y += 3;
                e.Graphics.DrawLine(new Pen(new SolidBrush(Color.Black)), new Point(start, _CurrentPoint.Y), new Point(start + len, _CurrentPoint.Y));
                _CurrentPoint.Y += 50;
                _CurrentPoint.X = _LeftMargin;
            }
        }

        #endregion

        #region 绘制统计图
        private void DrawImage(PrintPageEventArgs e)
        {
            //标题 居中
            _CurrentPoint.X = (e.PageBounds.Width) / 2 - _Width / 2;
            e.Graphics.DrawImage(GetPieImage(_ImgTitle, _DT), _CurrentPoint);
            _CurrentPoint.X = _LeftMargin;
            _CurrentPoint.Y += _Height + 50;
        }

        #endregion

        #region 绘制页尾
        private void DrawBottom(PrintPageEventArgs e)
        {
            if (_BottomStr != null)
            {
                int pageNumWidth = 70;
                int count = _BottomStr.Length;
                int width = (e.PageBounds.Width - _LeftMargin - _RightMargin - pageNumWidth) / (count + 1);
                int y = e.PageBounds.Height - _BottomMargin + 5;
                int x = _LeftMargin;
                //line
                e.Graphics.DrawLine(new Pen(new SolidBrush(Color.Black)), x, y, e.PageBounds.Width - _RightMargin, y);
                y += 5;
                for (int i = 0; i < count; i++)
                {
                    if (i > 0)
                        x += width;
                    e.Graphics.DrawString(_BottomStr[i], _ContentFont, new SolidBrush(Color.Black), x, y);
                }
                x = e.PageBounds.Width - _RightMargin - pageNumWidth;
                e.Graphics.DrawString(string.Format("第{0}页/共{1}页", _CurrentPageIndex, _PageCount), _ContentFont, new SolidBrush(Color.Black), x, y);
            }
        }

        #endregion

        #region 绘制表格和序号、数据

        private void DrawTableAndSerialNumAndData(PrintPageEventArgs e, int serialNumWidth, int colWidth)
        {
            int useAbleHeight = e.PageBounds.Height - _CurrentPoint.Y - _BottomMargin;
            int useAbleRowsCount = useAbleHeight / _RowHeight;
            int rowsCount = 0;
            if (_RowsCount - _CurrentRowsIndex > useAbleRowsCount)
                rowsCount = useAbleRowsCount;
            else
                rowsCount = _RowsCount - _CurrentRowsIndex;
            Point pp = new Point(_CurrentPoint.X, _CurrentPoint.Y);
            for (int i = 0; i <= rowsCount; i++)
            {
                e.Graphics.DrawLine(pen, _LeftMargin, _CurrentPoint.Y + i * _RowHeight, e.PageBounds.Width - _RightMargin, _CurrentPoint.Y + i * _RowHeight);
                //绘制数据
                if (i >= rowsCount)
                    break;
                DrawCellString((i + 1 + _CurrentRowsIndex).ToString(), pp, serialNumWidth, _ContentFont, e);
                pp.X += serialNumWidth;
                for (int j = 0; j < _DT.Columns.Count; j++)
                {
                    DrawCellString(_DT.Rows[i + _CurrentRowsIndex][j].ToString(), pp, colWidth, _ContentFont, e);
                    pp.X += colWidth;
                }
                pp.Y += _RowHeight;
                pp.X = _CurrentPoint.X;

            }
            //绘制竖线
            Point p = new Point(_CurrentPoint.X, _CurrentPoint.Y);
            e.Graphics.DrawLine(pen, p, new Point(p.X, p.Y + _RowHeight * rowsCount));
            p.X += serialNumWidth;
            e.Graphics.DrawLine(pen, p, new Point(p.X, p.Y + _RowHeight * rowsCount));
            for (int i = 1; i < _DT.Columns.Count; i++)
            {
                p.X += colWidth;
                e.Graphics.DrawLine(pen, p, new Point(p.X, p.Y + _RowHeight * rowsCount));
            }
            p.X = e.PageBounds.Width - _RightMargin;
            e.Graphics.DrawLine(pen, p, new Point(p.X, p.Y + _RowHeight * rowsCount));
            _CurrentRowsIndex += rowsCount;
        }

        #endregion

        #region 填充数据到单元格
        private void DrawCellString(string str, Point p, int colWidth, Font f, PrintPageEventArgs e)
        {
            int strWidth = (int)e.Graphics.MeasureString(str, f).Width;
            int strHeight = (int)e.Graphics.MeasureString(str, f).Height;
            p.X += (colWidth - strWidth) / 2;
            p.Y += 5;
            p.Y += (_RowHeight - strHeight) / 2;
            e.Graphics.DrawString(str, f, brush, p);
        }
        #endregion

        #region 绘制标题
        private void DrawTableHeader(PrintPageEventArgs e, int serialNumWidth, int colWidth)
        {
            //画框
            Point pp = new Point(_CurrentPoint.X, _CurrentPoint.Y);
            e.Graphics.DrawLine(pen, pp, new Point(e.PageBounds.Width - _RightMargin, pp.Y));
            pp.Y += _RowHeight;
            e.Graphics.DrawLine(pen, pp, new Point(e.PageBounds.Width - _RightMargin, pp.Y));
            pp = new Point(_CurrentPoint.X, _CurrentPoint.Y);
            e.Graphics.DrawLine(pen, pp, new Point(pp.X, pp.Y + _RowHeight));
            pp.X += serialNumWidth;
            e.Graphics.DrawLine(pen, pp, new Point(pp.X, pp.Y + _RowHeight));
            for (int i = 1; i < _DT.Columns.Count; i++)
            {
                pp.X += colWidth;
                e.Graphics.DrawLine(pen, pp, new Point(pp.X, pp.Y + _RowHeight));
            }
            pp.X = e.PageBounds.Width - _RightMargin;
            e.Graphics.DrawLine(pen, pp, new Point(pp.X, pp.Y + _RowHeight));
            //
            Point p = new Point(_CurrentPoint.X + 5, _CurrentPoint.Y);
            DrawCellString("序号", p, serialNumWidth, _ColumnsHeaderFont, e);
            p.X += serialNumWidth;
            int serial = 0;
            if (_ColumnsHeader.Length == _DT.Columns.Count + 1)
                serial = 1;
            for (int i = 0; i < _DT.Columns.Count; i++)
            {
                if (i != 0)
                    p.X += colWidth;
                DrawCellString(_ColumnsHeader[i + serial], p, colWidth, _ColumnsHeaderFont, e);
            }
            _CurrentPoint.X = _LeftMargin;
            _CurrentPoint.Y += _RowHeight;
        }
        #endregion

        #region 自定义设置打印预览对话框
        private void SetPrintPreviewDialog(PrintPreviewDialog pPreviewDialog)
        {
            System.Reflection.PropertyInfo[] pis = pPreviewDialog.GetType().GetProperties();
            for (int i = 0; i < pis.Length; i++)
            {
                if (pis[i].CanWrite)
                {
                    switch (pis[i].Name)
                    {
                        case "Dock":
                            pis[i].SetValue(pPreviewDialog, DockStyle.Fill, null);
                            break;
                        case "FormBorderStyle":
                            pis[i].SetValue(pPreviewDialog, FormBorderStyle.None, null);
                            break;
                        case "WindowState":
                            pis[i].SetValue(pPreviewDialog, FormWindowState.Normal, null);
                            break;
                        default:
                            break;
                    }
                }
            }
            #region 屏蔽默认的打印按钮，添加自定义的打印和保存按钮
            foreach (Control c in pPreviewDialog.Controls)
            {
                if (c is ToolStrip)
                {
                    ToolStrip ts = (ToolStrip)c;
                    ts.Items[0].Visible = false;
                    //print
                    ToolStripButton toolStripBtn_Print = new ToolStripButton();
                    toolStripBtn_Print.Text = "打印报表";
                    toolStripBtn_Print.ToolTipText = "打印当前报表";
                    toolStripBtn_Print.Image = Properties.Resources.Printer.ToBitmap();
                    toolStripBtn_Print.Click +=
delegate(object sender, EventArgs e)
{
    PrintDialog pd = new PrintDialog();
    pd.Document = pPreviewDialog.Document;
    pd.UseEXDialog = true;
    if (pd.ShowDialog() == DialogResult.OK)
        pPreviewDialog.Document.Print();
};
                    ToolStripButton toolStripBtn_SaveAsExcel = new ToolStripButton();
                    toolStripBtn_SaveAsExcel.Text = "导出到Excel";
                    toolStripBtn_SaveAsExcel.ToolTipText = "导出报表到Excel";
                    toolStripBtn_SaveAsExcel.Image = Properties.Resources.Excel.ToBitmap();
                    toolStripBtn_SaveAsExcel.Click +=
delegate(object sender, EventArgs e)
{
    SaveFileDialog f = new SaveFileDialog();
    f.Title = "数据保存到Excel...";
    f.Filter = "Excel文档(*.xls)|*.xls";
    if (f.ShowDialog() == DialogResult.OK)
    {
        Exception ex = SaveAsExcel(f.FileName);
        if (ex == null)
            MessageBox.Show("导出成功！");
        else
            MessageBox.Show("导出失败：" + ex.Message);
    }
    f.Dispose();
};
                    ToolStripSeparator tss = new ToolStripSeparator();
                    ts.Items.Insert(0, toolStripBtn_Print);
                    ts.Items.Insert(1, toolStripBtn_SaveAsExcel);
                    ts.Items.Insert(2, tss);
                }
            }
            #endregion
        }
        #endregion

        #endregion
    }
    /// <summary>
    /// 报表(导出)参数
    /// </summary>
    [Serializable]
    public class PrintArgs
    {
        string title;

        public string Title
        {
            get { return title; }
            set { title = value; }
        }
        string imgTitle;

        public string ImgTitle
        {
            get { return imgTitle; }
            set { imgTitle = value; }
        }
        string[] columnsHeader;

        public string[] ColumnsHeader
        {
            get { return columnsHeader; }
            set { columnsHeader = value; }
        }
        string[] bottomStr;

        public string[] BottomStr
        {
            get { return bottomStr; }
            set { bottomStr = value; }
        }

        public PrintArgs(string title, string imgTitle, string[] columnsHeader, string[] bottomStr)
        {
            this.title = title;
            this.imgTitle = imgTitle;
            this.columnsHeader = columnsHeader;
            this.bottomStr = bottomStr;
        }
    }
}