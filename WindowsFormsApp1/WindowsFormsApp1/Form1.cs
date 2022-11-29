using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "TEST.xlsm");
            var workbook = new XLWorkbook(filePath);
            IXLWorksheet ws = workbook.Worksheet("振り分け購買依頼");

            ////仕入先
            //DataTable dt = new DataTable();
            //dt.Columns.Add("仕入先コード", typeof(string));
            //dt.Columns.Add("仕入先名称", typeof(string));
            //for (int rowIndex = 1;rowIndex <= 10000;rowIndex++)
            //{
            //    DataRow dr = dt.NewRow();
            //    dr["仕入先コード"] = rowIndex;
            //    dr["仕入先名称"] = "名称"+rowIndex;
            //    dt.Rows.Add(dr);
            //}
            //workbook.Worksheet("仕入先").Cell("A1").InsertTable(dt);

            ////仕入先
            //DataTable dt2 = new DataTable();
            //dt2.Columns.Add("品目コード", typeof(string));
            //dt2.Columns.Add("プラント", typeof(string));
            //dt2.Columns.Add("所要日", typeof(DateTime));
            //dt2.Columns.Add("標準仕入先コード", typeof(string));
            //for (int rowIndex = 1; rowIndex <= 10000; rowIndex++)
            //{
            //    DataRow dr = dt2.NewRow();
            //    dr["品目コード"] = rowIndex;
            //    dr["プラント"] = "PW20";
            //    dr["所要日"] = new DateTime(2022, 11, 1);
            //    dr["標準仕入先コード"] = rowIndex;
            //    dt2.Rows.Add(dr);
            //}
            //workbook.Worksheet("SAP標準仕入れ先").Cell("A1").InsertTable(dt2);

            //var rowRange = ws.Row(3);
            //for (int rowIndex = 1; rowIndex <= 10000; rowIndex++)
            //{
            //    rowRange.CopyTo(ws.Row(rowIndex + 3));
            //}

            var rowRange = ws.Range(3, 1, 3, 15);
            for (int rowIndex = 1; rowIndex <= 10000; rowIndex++)
            {
                rowRange.CopyTo(ws.Cell(rowIndex + 3, 1));
            }

            for (int rowIndex = 1; rowIndex <= 10000; rowIndex++)
            {
                var valueArray = new[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15" };
                ws.Cell(rowIndex + 2, 1).InsertData(new[] { valueArray });

                //////値セット
                ////if (rowIndex > 1)
                ////{
                ////    //コピー
                ////    IXLRange range = ws.Range(3, 5, 3, 5);
                ////    ////ペースト
                ////    ws.Cell(rowIndex + 2, 7).Value = range;
                ////}

                //IXLCell cell = ws.Cell(rowIndex + 2, 5);
                //cell.Value = "1";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Protection.Locked = false;

                //cell = ws.Cell(rowIndex + 2, 6);
                //cell.Value = "2";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                ////if (rowIndex > 1)
                ////{
                ////    //コピー
                ////    IXLRange range = ws.Range(3, 7, 3, 7);
                ////    ////ペースト
                ////    ws.Cell(rowIndex + 2, 7).Value = range;
                ////}

                //cell = ws.Cell(rowIndex + 2, 7);
                //cell.Value = "3";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Protection.Locked = false;

                //cell = ws.Cell(rowIndex + 2, 8);
                //cell.Value = "4";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;


                //////値セット
                ////if (rowIndex > 1)
                ////{
                ////    //コピー
                ////    IXLRange range = ws.Range(3, 9, 3, 9);
                ////    ////ペースト
                ////    ws.Cell(rowIndex + 2, 9).Value = range;
                ////}
                //cell = ws.Cell(rowIndex + 2, 9);
                //cell.Value = "5";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                ////cell = ws.Cell(rowIndex + 2, 10);

                //cell.Value = "6";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                //cell = ws.Cell(rowIndex + 2, 11);
                //cell.Value = "7";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                //cell = ws.Cell(rowIndex + 2, 12);
                //cell.Value = "8";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                //cell = ws.Cell(rowIndex + 2, 13);
                //cell.Value = "9";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                //cell = ws.Cell(rowIndex + 2, 14);
                //cell.Value = "10";
                ////cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ////cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            }

            ////コピー
            //IXLRange range = ws.Range(3, 1, 3, 15);
            ////ペースト
            //ws.Cell(4, 1).Value = range;
            ////値セット
            //ws.Cell(4, 5).Style.NumberFormat.Format = "@";
            //ws.Cell(4, 5).Value = "2";

            //保存
            workbook.SaveAs(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "TEST_NEW.xlsm"));

            MessageBox.Show("完了");
        }
    }
}
