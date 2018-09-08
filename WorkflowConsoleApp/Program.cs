using System;
using System.Linq;
using System.Activities;
using System.Activities.Statements;

namespace WorkflowConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            /*BalaReva.Excel.Others.MergeCells MergeWorkFlow = new BalaReva.Excel.Others.MergeCells();
            MergeWorkFlow.FilePath = @"C:\TestProject\UiPathForm\ExcelComment\Data12.xlsx";
            MergeWorkFlow.SheetName = "Sheet1";
            MergeWorkFlow.CellText = "Harini Balamurugan";
            MergeWorkFlow.MergeRange = "G6:J6lo";

            WorkflowInvoker.Invoke(MergeWorkFlow);*/


            /*-------------------
             * BalaReva.Excel.Others.DeleteSheet delSheet = new BalaReva.Excel.Others.DeleteSheet();
            delSheet.FilePath = @"C:\TestProject\UiPathForm\ExcelComment\Data1.xlsx";
            delSheet.SheetName = "BalaReva";

            WorkflowInvoker.Invoke(delSheet);*/

            /*BalaReva.PowerPoint.DeleteSlide deleteSlide = new BalaReva.PowerPoint.DeleteSlide();
            deleteSlide.FilePath = @"C:\TestProject\UiPathForm\PowerPoint\ex1.pptx";
            deleteSlide.SlideIndex = 1;
            WorkflowInvoker.Invoke(deleteSlide);*/


            Formater.DoubleFormater formatDbl = new Formater.DoubleFormater();
            formatDbl.AfterDecimalFormat = "";
            formatDbl.InputValue = 1212551.55;
            formatDbl.Format= "#,###,##";

            WorkflowInvoker.Invoke(formatDbl);
        }
    }
}
