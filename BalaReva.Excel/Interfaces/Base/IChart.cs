namespace BalaReva.Excel.Interfaces
{
    using BalaReva.Excel.Charts;
    using System.Activities;

    public interface IChart
    {
        InArgument<string> FilePath { get; set; }
        InArgument<string> SheetName { get; set; }
        InArgument<string> CellRange { get; set; }
        InArgument<string> ChartTitle { get; set; }
        ChartSize Size { get; set; }
        ShowOptions Options { get; set; }
    }
}
