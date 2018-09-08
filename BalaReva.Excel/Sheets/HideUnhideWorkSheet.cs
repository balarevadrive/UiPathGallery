namespace BalaReva.Excel.Sheets
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("HideUnhide")]
    [Designer(typeof(ExcelSelection))]
    public class HideUnhideWorkSheet : BaseExcelNew
    {
        [Category("Input")]
        [Description("Sheet Visibility")]
        [DisplayName("Sheet Visibility")]
        public XlSheetVisibility SheetVisibility { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            this.DoCopyToWorkBook();
        }

        private void DoCopyToWorkBook()
        {
            try
            {
                base.InitWorkSheet();

                if (base.xlWorkSheet != null)
                {
                    if (SheetVisibility == XlSheetVisibility.Hidden)
                    { 
                        base.xlWorkSheet.Visible = ExcelObj.XlSheetVisibility.xlSheetHidden;
                    }
                    else
                    {
                        base.xlWorkSheet.Visible = ExcelObj.XlSheetVisibility.xlSheetVisible;
                    }

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
