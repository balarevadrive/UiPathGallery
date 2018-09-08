namespace BalaReva.Excel.Interfaces
{
    using BalaReva.Excel.Charts;

    public interface IColumnChart
    {
        ColumnChartEnum ChartType { get; set; }
    }
}
