namespace BalaReva.Excel.Sheets
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Delete Sheet")]
    [Designer(typeof(ExcelSelection))]
    public class DeleteSheet : BaseExcelNew
    {
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);

                this.DoDeleteSheet();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoDeleteSheet()
        {
            try
            {
                base.InitWorkSheet();

                if (xlWorkSheet != null)
                {
                    ExcelObj.Sheets sheets= xlWorkBook.Worksheets;

                    sheets[SheetIndex].Delete();

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
