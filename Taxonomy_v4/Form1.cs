using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;

namespace Taxonomy_v4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Bt_openExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "(.xlsx)|*.xlsx"
            };
            openfile.ShowDialog();

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook = excelApp.Workbooks.Open(
                openfile.FileName.ToString(), 0, true, 5, "", "", true,
                XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet excelSheet =
                (Worksheet)excelBook.Worksheets.get_Item(1);
            Range excelRange = excelSheet.UsedRange;
            string strCellData = "";
            int RowIndex = excelApp.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
            int ColumnIndex = excelApp.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;
            System.Data.DataTable dt = new System.Data.DataTable();
            for (ColumnIndex = 1; ColumnIndex <= excelRange.Columns.Count; ColumnIndex++)
            {
                string strColumn = "";
                var range = excelRange.Cells[1, ColumnIndex] as Range;
                if (range != null)
                    if (range.Value2 != null) strColumn = range.Value2.ToString();
                dt.Columns.Add(strColumn, typeof(string));
            }
            for (RowIndex = 1; RowIndex <= excelRange.Rows.Count; RowIndex++)
            {
                string strData = "";
                for (ColumnIndex = 1; ColumnIndex <= excelRange.Columns.Count; ColumnIndex++)
                {
                    var range = excelRange.Cells[RowIndex, ColumnIndex] as Range;
                    if (range != null)
                    {
                        strCellData = range.Value2 != null ? range.Value2.ToString() : String.Empty;
                        strData += strCellData + "|";
                    }

                }
                strData = strData.Remove(strData.Length - 1, 1);

                if (!string.IsNullOrWhiteSpace(strData)) dt.Rows.Add(strData.Split('|'));
            }
            dtGrid.DataSource = dt.DefaultView;
            dt.Rows[0].Delete();
            excelBook.Close(true, null, null);
            excelApp.Quit();
        }

        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void НормировкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var list = new List<double>();
            int rowcnt = dtGrid.Rows.Count;
            for (int j = 0; j < dtGrid.Columns.Count; j++)
            {
                for (int i = 0; i < rowcnt; i++)
                {
                    list.Add(Convert.ToDouble(dtGrid[j, i].Value));
                }
                var max = list.Max();
                var min = list.Min();
                var listAfterNorm = new List<double>();
                if (max == min)
                    listAfterNorm.AddRange(Enumerable.Range(0, list.Count).Select(_ => min == 0 ? 0.0 : 1.0));
                else
                    listAfterNorm.AddRange(list.Select(t => Math.Round(((t - min) / (max - min)), 3)));
                for (int i = 0; i < listAfterNorm.Count; i++)
                {
                    dtGrid[j, i].Value = listAfterNorm[i].ToString();
                }
                list.Clear();
                listAfterNorm.Clear();

            }
        }

        private void ТаксономияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double porog = Convert.ToDouble(Interaction.InputBox("Введите порог", "Ввод порога", "0,5", 100, 100));
            int num_elm_tax = 0;
            List<double> list_taxon1 = new List<double>();
            List<double> list_taxon2 = new List<double>();
            int taxonCnt = 1;
            richTextBox1.Text = "Таксон " + taxonCnt + ":";
            for (int j = 1; j < dtGrid.ColumnCount - 1; j++)
            {

                for (int i = 0; i < dtGrid.RowCount - 1; i++)
                    if (j + 1 != dtGrid.ColumnCount)
                    {
                        if (dtGrid.Rows[i].Cells[j].Value != null)
                            list_taxon1.Add(double.Parse(dtGrid.Rows[i].Cells[j].Value.ToString()));
                        if (dtGrid.Rows[i].Cells[j + 1].Value != null)
                            list_taxon2.Add(double.Parse(dtGrid.Rows[i].Cells[j + 1].Value.ToString()));
                    }


                if (GetRadius(list_taxon1, list_taxon2) <= porog)
                {
                    num_elm_tax++;
                    richTextBox1.AppendText(dtGrid.Columns[j + 1].Name.ToString() + ",");
                }
                else
                {
                    taxonCnt++;
                    num_elm_tax = 0;
                    richTextBox1.AppendText(dtGrid.Columns[j + 1].Name.ToString() + Environment.NewLine + "Таксон " + taxonCnt + ":");
                }
                list_taxon1.Clear();
                list_taxon2.Clear();
            }

        }

        private double GetRadius(List<double> l1, List<double> l2)
        {
            double radius = 0;
            for (int i = 0; i < l1.Count; i++)
                radius = radius + Math.Pow((l1[i] - l2[i]), 2);
            return Math.Sqrt(radius);
        }
    }
}
