namespace BalaReva.Excel.Others
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Rename Work Sheet")]
    [Designer(typeof(RenameSheetDesigner))]
    public class RenameSheet : BaseExcelNew
    {

        [Category("Input")]
        [RequiredArgument]
        [Description("New name for the sheet")]
        [DisplayName("New Sheet Name")]
        public InArgument<string> NewSheetName { get; set; }

        private string strNewSheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            this.strNewSheetName = NewSheetName.Get(context);

            this.DoRename();
        }

        private void DoRename()
        {
            try
            {
                base.InitWorkSheet();

                if (xlWorkSheet != null)
                {
                    // initialize the work sheet object
                    xlWorkSheet.Name = this.strNewSheetName;

                    base.SaveWorkBook(true);
                }
                else
                {
                    base.ClearObject();
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

