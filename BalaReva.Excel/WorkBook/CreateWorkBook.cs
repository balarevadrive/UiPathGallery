namespace BalaReva.Excel.WorkBook
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Create WorkBook")]
    public class CreateWorkBook : BaseExcelNew
    {
        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            this.DoCreateWorkBook();
        }

        private void DoCreateWorkBook()
        {
            try
            {
                base.InitApp();

                base.xlWorkBook = base.xlApp.Workbooks.Add();
                base.IsWorkBookOpened = true;
                base.xlWorkSheet = (ExcelObj.Worksheet)base.xlWorkBook.ActiveSheet;
                base.xlWorkSheet.Name = base.strSheetName;
                base.IsWorkSheetOpened = true;

                if (!string.IsNullOrEmpty(base.strPassword))
                {
                    xlWorkBook.Password = base.strPassword;
                }

                xlWorkBook.SaveAs(strFilePath);

                base.SaveWorkBook(false);
            }
            catch (Exception ex)
            {
                base.ClearObject();
                throw ex;
            }
        }
    }
}
