using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            int A = 0, B = 0, C = 0;
            Excel.Application excel = new Excel.Application();
            Object t = Type.Missing;
            excel.Visible = true;
            excel.SheetsInNewWorkbook = 3;
            excel.Workbooks.Add(Type.Missing);
            var book = excel.Workbooks[1];
            var sheets = excel.Worksheets;
            var worksheet = (Excel.Worksheet)sheets.get_Item(1);
            var excelcell = worksheet.Cells;
            excelcell = worksheet.get_Range("A1", "E7");
            excelcell.Borders.ColorIndex = 1;
            excelcell.Borders.Weight = 3;
            worksheet.get_Range("A1", t).Value2 = "Регион";
            worksheet.get_Range("B1", t).Value2 = "Лыжи";
            worksheet.get_Range("C1", t).Value2 = "Коньки";
            worksheet.get_Range("D1", t).Value2 = "Санки";
            worksheet.get_Range("E1", t).Value2 = "Всего";

            int rowNumber = 2;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Index == dataGridView1.Rows.Count - 1)
                {
                    continue;
                }

                worksheet.get_Range("A" + rowNumber, t).Value2 = row.Cells[0].Value;
                worksheet.get_Range("B" + rowNumber, t).Value2 = row.Cells[1].Value;
                worksheet.get_Range("C" + rowNumber, t).Value2 = row.Cells[2].Value;
                worksheet.get_Range("D" + rowNumber, t).Value2 = row.Cells[3].Value;

                var a = Int32.Parse(row.Cells[1].Value.ToString());
                var b = Int32.Parse(row.Cells[2].Value.ToString());
                var c = Int32.Parse(row.Cells[3].Value.ToString());

                A += a;
                B += b;
                C += c;

                worksheet.get_Range("E" + rowNumber, t).Value2 = a + b + c;
                rowNumber++;
                if (rowNumber == 8)
                {
                    break;
                }
            }

            worksheet.get_Range("A" + rowNumber, t).Value2 = "Среднее";
            worksheet.get_Range("B" + rowNumber, t).Value2 = Math.Round((decimal)A / (dataGridView1.Rows.Count - 1));
            worksheet.get_Range("C" + rowNumber, t).Value2 = Math.Round((decimal)B / (dataGridView1.Rows.Count - 1));
            worksheet.get_Range("D" + rowNumber, t).Value2 = Math.Round((decimal)C / (dataGridView1.Rows.Count - 1));

            Chart graph = (Excel.Chart)excel.Charts.Add(t, t, t, t);
            // graph.SetSourceData(worksheet.get_Range("E2", "E6"), XlRowCol.xlColumns);
            // graph.ChartType = XlChartType.xlColumnClustered;
            graph.HasLegend = false;
            graph.HasTitle = true;
            graph.ChartTitle.Caption = "Ведомость";
            Axis hAxis = (Excel.Axis)graph.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            hAxis.HasTitle = false;
            Axis vAxis = (Excel.Axis)graph.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            vAxis.HasTitle = false;

            Object FileName = @"D:\Учет.xls";
            Object FileFormat = Type.Missing;

            Object Password = Type.Missing;

            Object WriteRes = Type.Missing;
            Object ReadOnlyRecommended = Type.Missing;
            Object CreateBac = Type.Missing;
            Object Confkict = Type.Missing;
            Object AddTo = Type.Missing;
            Object TextCode = Type.Missing;
            Object TextVisual = Type.Missing;
            Object Local = Type.Missing;

            excel.ActiveWorkbook.SaveAs(
                FileName, FileFormat,
                Password, WriteRes, ReadOnlyRecommended,
                CreateBac, XlSaveAsAccessMode.xlNoChange, Confkict, AddTo, TextCode,
                TextVisual, Local);


            excel.Quit();

        }
    }

}
