namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.IO;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    public abstract class BaseChart : BaseExcelNew, IChart
    {


        [Category("Input")]
        [RequiredArgument]
        [Description("Cell Range")]
        [DisplayName("Cell Range")]
        public InArgument<string> CellRange { get; set; }


        [Category("Input")]
        [Description("Chart Title")]
        [DisplayName("Chart Titlte")]
        public InArgument<string> ChartTitle { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public ChartSize Size { get; set; } = new ChartSize();

        [Category("Input")]
        public ShowOptions Options { get; set; } = new ShowOptions();

        [Category("Input")]
        [Description("Full file path")]
        [DisplayName("Image Copy")]
        public InArgument<string> ImageCopy { get; set; }

        private string strCellRange { get; set; }
        private double dblChartLeft { get; set; }
        private double dblChartTop { get; set; }
        private double dblChartWidth { get; set; }
        private double dblChartHeight { get; set; }
        private string strChartTitle { get; set; }
        private string strImagecopy { get; set; }


        // Initalized the context values to local variables
        protected void InitValue(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);

                this.strCellRange = CellRange.Get(context);

                this.dblChartLeft = Size.Left;
                this.dblChartTop = Size.Top;
                this.dblChartWidth = Size.Width;
                this.dblChartHeight = Size.Height;
                this.strChartTitle = ChartTitle.Get(context);
                this.strImagecopy = ImageCopy.Get(context);

            }
            catch (Exception ex)
            {
                throw new Exception("InitValue" + ex.Message);
            }
        }

        protected void DrawChart(XlChartType ChartType)
        {
            // Variable declaration 

            object misValue = System.Reflection.Missing.Value;
            ExcelObj.Range chartRange=null;

            try
            {
                this.Validate();

                // split the cell range for the work sheet range
                string cell1 = strCellRange.Split(':')[0];
                string cell2 = strCellRange.Split(':')[1];

                base.InitWorkSheet();


                if (base.SheetIndex > 0 && base.xlWorkSheet != null)
                {
                    ExcelObj.ChartObjects xlCharts = (ExcelObj.ChartObjects)base.xlWorkSheet.ChartObjects(Type.Missing);

                    try
                    {
                        // Define the cell range to the work sheet 
                        chartRange = xlWorkSheet.get_Range(cell1, cell2);
                    }
                    catch (Exception)
                    {
                        base.ClearObject();

                        throw new Exception("Invalid Range");
                    }

                    // fix chart position by left,top,width,height
                    ExcelObj.ChartObject chartObj = (ExcelObj.ChartObject)xlCharts.Add(dblChartLeft, dblChartTop, dblChartWidth, dblChartHeight);

                    // Define char object
                    ExcelObj._Chart chart = chartObj.Chart;

                    // assign the datasource for the chart
                    chart.SetSourceData(chartRange, misValue);

                    // set the chart type.
                    chart.ChartType = ChartType;

                    // Ledgend show / hide
                    chart.HasLegend = this.Options.ShowLedgend;

                    // assign the chart title based on empty 
                    if (!string.IsNullOrEmpty(strChartTitle))
                    {
                        chart.HasTitle = true;
                        chart.ChartTitle.Text = strChartTitle;
                    }
                    else
                    {
                        chart.HasTitle = false;
                    }

                    XlDataLabelsType XlDataLabelsType = (XlDataLabelsType)this.Options.DataLabelsType;

                    XlDataLabelsType XlLedgendKey = (XlDataLabelsType)this.Options.LedgendKey;

                    // assign the data lable properties
                    chart.ApplyDataLabels(XlDataLabelsType, XlLedgendKey, this.Options.AutoText, this.Options.HasLeaderLines,
                        this.Options.ShowSeriesName, this.Options.ShowCategoryName, this.Options.ShowValue,
                        this.Options.ShowPercentage, this.Options.ShowBubbleSize, this.Options.Separator);

                    if (this.strImagecopy != null && !string.IsNullOrEmpty(this.strImagecopy))
                    {
                        string strExten = new FileInfo(this.strImagecopy).Extension;
                        strExten = strExten.Replace(".", "");
                        chart.Export(this.strImagecopy, strExten.ToUpper(), misValue);
                    }

                    base.SaveWorkBook(true);
                }
            }
            catch (ValidationException Vex)
            {
                base.ClearObject();
                throw Vex;
            }
            catch (Exception ex)
            {
                base.ClearObject();

                throw new Exception("DrawChart : " + ex.Message);
            }
        }


        private void Validate()
        {
            if (!File.Exists(this.strFilePath))
            {
                throw new Exception("Invalid Excel file");
            }
        }
    }
}
