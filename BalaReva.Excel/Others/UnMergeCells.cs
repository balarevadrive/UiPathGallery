namespace BalaReva.Excel.Others
{
    using BalaReva.Excel.Design;
    using Enums;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("UnMerge Cells")]
    [Designer(typeof(ExcelSelection))]
    public class UnMergeCells : BaseExcelNew
    {
        [Category("Input"), RequiredArgument]
        [Description("Example A1:D10")]
        [DisplayName("UnMerge Range")]
        public InArgument<string> UnMergeRange { get; set; }

        [Category("Alignment"), Description("Vertical Alignment")]
        [DisplayName("Vertical Alignment")]
        public AlignmentEnum VerticalAlignment { get; set; } = AlignmentEnum.Center;

        [Category("Alignment"), Description("Horizontal Alignment")]
        [DisplayName("Horizontal Alignment")]
        public AlignmentEnum HorizontalAlignment { get; set; } = AlignmentEnum.Center;

        private string strMergeRange { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);


                strMergeRange = UnMergeRange.Get(context);

                this.UnMergeXlCells();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void UnMergeXlCells()
        {
            try
            {
                base.InitWorkSheet();

                ExcelObj.Range xlRange = base.GetSheetRange(strMergeRange);

                if (xlRange != null)
                {
                    xlRange.UnMerge();
                    xlRange.VerticalAlignment = VerticalAlignment;
                    xlRange.HorizontalAlignment = HorizontalAlignment;

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
                throw ex;
            }
        }
    }
}
