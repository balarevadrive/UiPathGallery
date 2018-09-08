namespace BalaReva.Excel.Charts
{
    using BalaReva.Excel.Design;
    using BalaReva.Excel.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Column Chart")]
    [Designer(typeof(ColumnChartDesigner))]
    public class ColumnChart : BaseChart, IColumnChart
    {
        [Category("Input")]
        [DisplayName("Chart Type")]
        public ColumnChartEnum ChartType { get; set; } = ColumnChartEnum.ColumnClustered;

        protected override void Execute(CodeActivityContext context)
        {
            // Initalize the values to base class variables
            base.InitValue(context);

            base.DrawChart((XlChartType)ChartType);
        }
    }
}
