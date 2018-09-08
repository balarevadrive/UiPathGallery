namespace BalaReva.Excel.Interfaces
{
    using BalaReva.Excel.Charts;

    public interface ILineChart
    {
        LineChartEnum ChartType { get; set; }
    }
}
