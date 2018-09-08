namespace BalaReva.Excel.Sheets
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Protect UnProtect")]
    [Designer(typeof(ExcelSelection))]
    public class ProtectUnProtectSheet : BaseExcelNew
    {
        [Category("Input")]
        [Description("ProtectPassword")]
        [DisplayName("Protect Password")]
        public InArgument<string> ProtectPassword { get; set; }

        [Category("Input")]
        [Description("Protect Type")]
        [DisplayName("Protect Type"),DefaultValue(ProtectUnProtectEnum.Protect)]
        public ProtectUnProtectEnum ProtectType { get; set; } = ProtectUnProtectEnum.Protect;

        private string strProtectPwd { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            strProtectPwd = ProtectPassword.Get(context);

            this.DoProtect();
        }

        private void DoProtect()
        {
            try
            {
                base.InitWorkSheet();

                if (base.xlWorkSheet != null)
                {
                    if (ProtectType == ProtectUnProtectEnum.Protect)
                    {
                        base.xlWorkSheet.Protect(strProtectPwd);
                    }
                    else
                    {
                        base.xlWorkSheet.Unprotect(strProtectPwd);
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
