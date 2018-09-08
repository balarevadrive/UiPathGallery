namespace BalaReva.Excel.Others
{
    using BalaReva.Excel.Design;
    using BalaReva.Excel.Enums;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using System.Windows.Forms;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Merge Cells")]
    [Designer(typeof(ExcelSelection))]
    public class MergeCells : BaseExcelNew
    {
        [Category("Input"), RequiredArgument]
        [Description("Example A1:D10"), DisplayName("Merge Range")]
        public InArgument<string> MergeRange { get; set; }

        [Category("Input"), RequiredArgument]
        [Description("Cell Text"), DisplayName("Cell Text")]
        public InArgument<string> CellText { get; set; }

        [Category("Alignment"), Description("Vertical Alignment")]
        [DisplayName("Vertical Alignment")]
        public AlignmentEnum VerticalAlignment { get; set; } = AlignmentEnum.Center;

        [Category("Alignment"), Description("Horizontal Alignment")]
        [DisplayName("Horizontal Alignment")]
        public AlignmentEnum HorizontalAlignment { get; set; } = AlignmentEnum.Center;

        private string strMergeRange { get; set; }
        private string strCellText { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadVariables(context);

                strMergeRange = MergeRange.Get(context);
                strCellText = CellText.Get(context);

                this.MergeXlCells();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void MergeXlCells()
        {
            try
            {
                base.InitWorkSheet();

                ExcelObj.Range xlRange = base.GetSheetRange(strMergeRange);

                if (xlRange != null)
                {
                    xlRange.Merge();
                    xlRange.Value = strCellText;

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
                base.ClearObject();

                MessageBox.Show(ex.Message);
            }
        }
    }
}
