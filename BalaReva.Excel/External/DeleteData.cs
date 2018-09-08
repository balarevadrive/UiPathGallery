namespace BalaReva.Excel.External
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Delete Data")]
    [Designer(typeof(DeleteDataDesign))]
    public class DeleteData : BaseExcelNew
    {

        [Category("Input")]
        [RequiredArgument]
        [Description("Example A1:D10")]
        [DisplayName("Delete Range")]
        public InArgument<string> DeleteRange { get; set; }


        private string strDeleteRange { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);

                strDeleteRange = DeleteRange.Get(context);

                this.DeleteXlData();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DeleteXlData()
        {
            try
            {
                base.InitWorkSheet();

                ExcelObj.Range xlRange = base.GetSheetRange(strDeleteRange);

                if (xlRange != null)
                {
                    xlRange.Clear();
                    base.SaveWorkBook(true);
                }
                else
                {
                    base.ClearObject();
                    throw new Exception("Invalid Range");
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
