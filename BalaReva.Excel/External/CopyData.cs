namespace BalaReva.Excel.External
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Copy Data")]
    [Designer(typeof(CopyDataDesign))]
    public class CopyData : BaseExcelNew
    {

        [Category("Input")]
        [RequiredArgument]
        [Description("Example A1:D10")]
        [DisplayName("Copy Range")]
        public InArgument<string> CopyRange { get; set; }


        private string strCopyRange { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);
                strCopyRange = CopyRange.Get(context);

                this.CopyXlData();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void CopyXlData()
        {
            try
            {
                base.InitWorkSheet();

                ExcelObj.Range xlRange = base.GetSheetRange(strCopyRange);
                if (xlRange != null)
                {
                    xlRange.Copy();

                    base.SaveWorkBook(false);
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
