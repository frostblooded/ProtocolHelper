using Microsoft.Office.Interop.Word;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ProtocolHelper
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Word.Application app;
        Document doc;

        List<System.Windows.Forms.GroupBox> groupBoxes;
        int currentGroupBox;

        public Form1()
        {
            InitializeComponent();
            this.Text = "Protocol Helper";
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MinimizeBox = false;
            this.MaximizeBox = false;

            InitializeMicrosoftWord();
            CreateGroupBoxes();

            currentGroupBox = 0;
            ShowGroupBox(currentGroupBox);
        }

        private void CreateGroupBoxes()
        {
            groupBoxes = new List<System.Windows.Forms.GroupBox>();

            //First group box
            System.Windows.Forms.GroupBox groupBox1 = new System.Windows.Forms.GroupBox();
            groupBox1.Width = this.Width;
            groupBox1.Height = this.Height;

            System.Windows.Forms.Button createFileButton = new System.Windows.Forms.Button();
            createFileButton.Text = "Create New File";
            createFileButton.Click += CreateDocumentButtonClick;
            createFileButton.Width = 90;
            createFileButton.Location = new System.Drawing.Point(this.ClientSize.Width / 2 - createFileButton.Width / 2, 40);

            System.Windows.Forms.Button openFileButton = new System.Windows.Forms.Button();
            openFileButton.Text = "Open Existing File";
            openFileButton.Click += OpenDocumentButtonClick;
            openFileButton.Width = 100;
            openFileButton.Location = new System.Drawing.Point(this.ClientSize.Width / 2 - openFileButton.Width / 2, 70);

            groupBox1.Controls.Add(createFileButton);
            groupBox1.Controls.Add(openFileButton);

            groupBoxes.Add(groupBox1);

            //Second group box
            System.Windows.Forms.GroupBox groupBox2 = new System.Windows.Forms.GroupBox();
            groupBox2.Width = this.Width;
            groupBox2.Height = this.Height;

            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            label.Text = "When you click the button a table will be created.\nPlease insert the names of the values in the blue column\nand in the others insert the values.\n\nRows will always be 2\nHow much columns do you want?";
            label.Location = new System.Drawing.Point(5, 5);
            label.AutoSize = true;

            TextBox columnsInput = new TextBox();
            columnsInput.Location = new System.Drawing.Point(this.ClientSize.Width / 2 - columnsInput.Width / 2, 100);

            System.Windows.Forms.Button createTableButton = new System.Windows.Forms.Button();
            createTableButton.Text = "Create Table";
            createTableButton.Click += CreateTableButtonClick;
            createTableButton.Width = 80;
            createTableButton.Location = new System.Drawing.Point(this.ClientSize.Width / 2 - createTableButton.Width / 2, 130);

            groupBox2.Controls.Add(label);
            groupBox2.Controls.Add(columnsInput);
            groupBox2.Controls.Add(createTableButton);

            groupBoxes.Add(groupBox2);

            //Third group box
            System.Windows.Forms.GroupBox groupBox3 = new System.Windows.Forms.GroupBox();
            groupBox3.Width = this.Width;
            groupBox3.Height = this.Height;

            System.Windows.Forms.Button createChartButton = new System.Windows.Forms.Button();
            createChartButton.Text = "Create Chart";
            createChartButton.Click += CreateChartButtonClick;
            createChartButton.Location = new System.Drawing.Point(this.ClientSize.Width / 2 - createChartButton.Width / 2, this.ClientSize.Height / 2 - createChartButton.Height / 2);

            groupBox3.Controls.Add(createChartButton);

            groupBoxes.Add(groupBox3);
        }

        private void ShowGroupBox(int index)
        {
            this.Controls.Clear();
            this.Controls.Add(groupBoxes[index]);
        }

        private void InitializeMicrosoftWord()
        {
            app = new Microsoft.Office.Interop.Word.Application();
        }

        private void CreateDocumentButtonClick(object sender, EventArgs e)
        {
            doc = app.Documents.Add();

            app.Visible = true;
            ShowGroupBox(++currentGroupBox);
        }

        private void OpenDocumentButtonClick(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            string filePath = openFileDialog.FileName;
            doc = app.Documents.Open(filePath);

            app.Visible = true;
            ShowGroupBox(++currentGroupBox);
        }

        private void CreateTableButtonClick(object sender, EventArgs e)
        {
            //Get textbox
            var textBoxes = this.groupBoxes[currentGroupBox].Controls.OfType<TextBox>();
            var columnsTextBox = textBoxes.First<TextBox>();

            CreateTable(2, int.Parse(columnsTextBox.Text));
            InsertNewLine();
            ShowGroupBox(++currentGroupBox);
        }

        private void CreateTable(int rows, int columns)
        {
            //Insert table
            Microsoft.Office.Interop.Word.Range tableLocation = doc.Range(doc.Content.End - 1);
            Table table = doc.Tables.Add(tableLocation, rows, columns, WdDefaultTableBehavior.wdWord9TableBehavior);

            //Format table
            table.Range.set_Style("Table Grid 1");
            table.Range.ParagraphFormat.SpaceAfter = 6;
            table.Range.ParagraphFormat.SpaceBefore = 6;

            for (int i = 1; i <= 2; i++)
            {
                table.Cell(i, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorTurquoise;
            }

            Marshal.ReleaseComObject(tableLocation);
            Marshal.ReleaseComObject(table);
        }

        private void InsertNewLine()
        {
            doc.Range(doc.Content.End - 1).Text = Environment.NewLine;
        }


        private void CreateChartButtonClick(object sender, EventArgs e)
        {
            CreateChart();
        }

        private void CreateChart()
        {
            Microsoft.Office.Interop.Word.Range chartLocation = doc.Range(doc.Content.End - 1);
            InlineShape inlineShape = doc.InlineShapes.AddChart2(-1, Microsoft.Office.Core.XlChartType.xlXYScatterSmooth, chartLocation);

            var chart = inlineShape.Chart;
            var chartData = chart.ChartData;
            var chartWorkbook = chartData.Workbook;
            var chartWorksheet = chartWorkbook.ActiveSheet;

            var table = doc.Tables[doc.Tables.Count];

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    string tableCellData = table.Cell(i, j).Range.Text;
                    tableCellData = tableCellData.Replace("\r\a", String.Empty); //to remove some odd symbols in the end of cells
                    chartWorksheet.Cells[j, i] = tableCellData;
                }
            }

            var rangeBegin = chartWorksheet.Cells[1, 1];
            var rangeEnd = chartWorksheet.Cells[table.Columns.Count, table.Rows.Count];
            var chartRange = chartWorksheet.Range[rangeBegin, rangeEnd];

            chart.SetSourceData(chartWorksheet.Name + "!" + chartRange.Address, XlRowCol.xlColumns);

            Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(chartData);
            Marshal.ReleaseComObject(chartWorkbook);
            Marshal.ReleaseComObject(chartWorksheet);

            Marshal.ReleaseComObject(table);
            Marshal.ReleaseComObject(chartLocation);
            Marshal.ReleaseComObject(inlineShape);
        }
    }
}
