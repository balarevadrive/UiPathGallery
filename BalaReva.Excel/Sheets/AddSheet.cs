namespace BalaReva.Excel.Sheets
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Add Sheet")]
    [Designer(typeof(ExcelSelection))]
    public class AddSheet : BaseExcelNew
    {
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);

                this.DoAddSheet();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoAddSheet()
        {
            try
            {
                base.InitWorkBook();

                if (xlWorkBook!= null)
                {
                    ExcelObj._Worksheet  NewxlWorkSheet= xlWorkBook.Worksheets.Add();
                    NewxlWorkSheet.Name = strSheetName;

                    base.SaveWorkBook(true);
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
