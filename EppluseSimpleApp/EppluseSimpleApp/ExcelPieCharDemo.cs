using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;

namespace EppluseSimpleApp
{
    
    /// <summary>
    /// 饼图案例
    /// </summary>
    public class ExcelPieCharDemo
    {
        //空参构造
        public ExcelPieCharDemo()
        {
            PieChart();
        }

        public void PieChart() {
            //Excel路径
            string ExcelPath = AppDomain.CurrentDomain.BaseDirectory + "PieChart.xlsx";
            try
            {
                //检查文件是否存在
                if (File.Exists(ExcelPath))
                {
                    File.Delete(ExcelPath);
                }
            }
            catch{
                throw new Exception("文件删除异常，请确认文件是否被其他程序占用");                
            }
            //FileInfo对象
            FileInfo info = new FileInfo(ExcelPath);

            //使用Eppluse  需要引入 Eppluse.dll   using OfficeOpenXml;
            using (ExcelPackage package = new ExcelPackage(info)) {

                //获得workbook  对象
                ExcelWorkbook workbook = package.Workbook;
                //创建 sheet  
                ExcelWorksheet worksheet = workbook.Worksheets.Add("sheetName");

                //将DataTable资料写入Excel
                DataTable dt = GetData();
                worksheet.Cells[1, 1].LoadFromDataTable(dt,true);

                //获得sheet名
                string name=worksheet.Name;
                // cell中的值为某一块区域的最左上角单元格行列与最右下角单元格行列，也可以直接用单元格地址worksheet.Cells["A1:C3"]
                var total =worksheet.Names.Add("SubTotalName",worksheet.Cells[dt.Rows.Count+2,3,dt.Rows.Count+2,dt.Columns.Count]);
                //设置字体
                total.Style.Font.Italic = true;
                //设置公式
                total.Formula = $"SUM({worksheet.Cells[2,3].Address}:{worksheet.Cells[dt.Rows.Count+1,3].Address})";

                //设置某一区域的显示格式
                worksheet.Cells[2, 3, dt.Rows.Count + 1, dt.Columns.Count].Style.Numberformat.Format = "#,##0";

                //绘制饼图   using OfficeOpenXml.Drawing.Chart;
                var chart = worksheet.Drawings.AddChart("PieChart", eChartType.Pie3D) as ExcelPieChart;

                //设置饼图标题
                chart.Title.Text = "汇总分析";
                //设置饼图位置 参数说明：起始行 行偏移 起始列 列偏移
                chart.SetPosition(0,0,dt.Columns.Count,5);
                //设置饼图大小 宽度/高度
                chart.SetSize(600,300);

                //饼图取取值范围 相当于 worksheet.Cells[2,dt.Columns.Count,dt.Rows.Count+1,dt.Columns.Count]
                ExcelAddress valueAddress = new ExcelAddress(2,dt.Columns.Count,dt.Rows.Count+1,dt.Columns.Count);
                //设置饼图的取值，参数分别为 扇形图区域的值区域与该区域的名称的值的区域
                var ser = chart.Series.Add(valueAddress.Address, worksheet.Cells[2, 2, dt.Rows.Count + 1, 2].Address) as ExcelPieChartSerie;

                chart.DataLabel.ShowCategory = true;
                chart.DataLabel.ShowPercent = true;

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;

                //worksheet.View.PageLayoutView = false;

                //保存文件
                package.Save();
            }

        }

        //准备数据
        public DataTable GetData() {

            //准备数据表
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Product",typeof(string));//产品
            dt.Columns.Add("Quantity", typeof(int));//数量
            dt.Columns.Add("Price", typeof(double));//单价
            dt.Columns.Add("Value", typeof(double));//总计值

            //准备测试数据
            DataRow row1 = dt.NewRow();
            row1[0] = 1001;
            row1[1] = "HuaWei";
            row1[2] = 99;
            row1[3] = 4999;
            row1[4] = 99 * 4999;
            dt.Rows.Add(row1);

            row1 = dt.NewRow();
            row1[0] = 1002;
            row1[1] = "OPPO";
            row1[2] = 88;
            row1[3] = 1999;
            row1[4] = 88 * 1999;
            dt.Rows.Add(row1);

            row1 = dt.NewRow();
            row1[0] = 1003;
            row1[1] = "Iphone";
            row1[2] = 59;
            row1[3] = 5999;
            row1[4] = 59 * 5999;
            dt.Rows.Add(row1);

            row1 = dt.NewRow();
            row1[0] = 1004;
            row1[1] = "XiaoMi";
            row1[2] = 99;
            row1[3] = 3999;
            row1[4] = 99 * 3999;
            dt.Rows.Add(row1);

            row1 = dt.NewRow();
            row1[0] = 1005;
            row1[1] = "VIVO";
            row1[2] = 100;
            row1[3] = 2199;
            row1[4] = 100 * 2199;
            dt.Rows.Add(row1);

            row1 = dt.NewRow();
            row1[0] = 1006;
            row1[1] = "SAMSUNG";
            row1[2] = 99;
            row1[3] = 2599;
            row1[4] = 99 * 2599;
            dt.Rows.Add(row1);

            return dt;
        }
    }
}
