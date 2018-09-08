namespace BalaReva.Excel.Sheets
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;
    using ReleaseObj = System.Runtime.InteropServices;

    [DisplayName("Copy To WorkBook")]
    [Designer(typeof(ExcelSelection))]
    public class CopyToWorkBook : BaseExcelNew
    {
        [Category("Input")]
        [Description("New Sheet Name")]
        [DisplayName("New Sheet Name")]
        public InArgument<string> NewSheetName { get; set; }

        private string StrNewSheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);
            StrNewSheetName = NewSheetName.Get(context);

            this.DoCopyToWorkBook();
        }

        private void DoCopyToWorkBook()
        {
            try
            {
                base.InitWorkSheet();

                if (base.xlWorkSheet != null)
                {
                    ExcelObj.Worksheet NewSheet = (ExcelObj.Worksheet)base.xlWorkBook.Sheets.Item[base.SheetCount];
                    base.xlWorkSheet.Copy(Type.Missing, NewSheet);

                    if (StrNewSheetName != null && !string.IsNullOrEmpty(StrNewSheetName))
                    {
                        NewSheet = (ExcelObj.Worksheet)base.xlWorkBook.Sheets.Item[base.SheetCount];
                        NewSheet.Name = StrNewSheetName;
                    }

                    base.SaveWorkBook(true);

                    ReleaseObj.Marshal.ReleaseComObject(NewSheet);
                }
            }
            catch (Exception ex)
            {

                base.ClearObject();
                throw ex;
            }
        }
    }
}
