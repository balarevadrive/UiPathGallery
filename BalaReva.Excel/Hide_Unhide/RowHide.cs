namespace BalaReva.Excel.Hide_UnHide
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Row")]
    [Designer(typeof(RowHideDesigner))]
    public class RowHide : BaseExcelNew
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Like 1,2,3 ")]
        [DisplayName("Row Numbers")]
        public InArgument<int[]> RowNumbers { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Hide or Unhide")]
        [DisplayName("Hidden Type")]
        public HideEnum HiddenType { get; set; } = HideEnum.Hide;

        private int[] intRowNumber { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            this.intRowNumber = RowNumbers.Get(context);

            this.DoRowHide();
        }

        private void DoRowHide()
        {
            object misValue = System.Reflection.Missing.Value;

            try
            {
                base.InitWorkSheet();

                if (xlWorkSheet != null)
                {
                    foreach (Int16 item in this.intRowNumber)
                    {
                        try
                        {
                            xlWorkSheet.get_Range(item + ":" + item, misValue).EntireRow.Hidden = (HiddenType == HideEnum.Hide);
                        }
                        catch (Exception)
                        {
                            base.ClearObject();

                            throw new Exception("Invalid Range");
                        }
                    }

                    base.SaveWorkBook(true);
                }

            }
            catch (Exception ex)
            {
                base.ClearObject();
                throw new Exception("ColumnHide : " + ex.Message);
            }
        }
    }
}
