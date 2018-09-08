namespace BalaReva.Excel.Hide_UnHide
{
    using BalaReva.Excel.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Column")]
    [Designer(typeof(ColumnHideDesigner))]
    public class ColumnHide : BaseExcelNew
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Like A,B")]
        [DisplayName("Column Names")]
        public InArgument<string[]> ColumnNames { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Hide or Unhide")]
        [DisplayName("Hidden Type")]
        public HideEnum HiddenType { get; set; } = HideEnum.Hide;

        private string[] strColumns { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            this.strColumns = ColumnNames.Get(context);

            this.DoColumnHide();
        }

        private void DoColumnHide()
        {
            try
            {
                base.InitWorkSheet();

                if (xlWorkSheet != null)
                {
                    foreach (string item in this.strColumns)
                    {
                        try
                        {
                            xlWorkSheet.get_Range(item + ":" + item, misValue).EntireColumn.Hidden = (HiddenType == HideEnum.Hide);
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
