namespace BalaReva.Excel.Interfaces
{
    using System.Activities;

    internal interface IBaseCommnet
    {
        InArgument<string> FilePath { get; set; }
        InArgument<string> SheetName { get; set; }
        InArgument<string> Cell { get; set; }
    }
}
