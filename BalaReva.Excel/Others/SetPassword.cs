/*namespace BalaReva.Excel.Others
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    [DisplayName("Set Password")]
    [Designer(typeof(ExcelSelection))]
    public class SetPassword : BaseExcel
    {
        [Category("Input"), RequiredArgument]
        [Description("File path")]
        [DisplayName("File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [Description("File password")]
        [DisplayName("File Password")]
        public InArgument<string> FilePassword { get; set; }

        [Category("Input")]
        [Description("New password for the Excel Sheet")]
        [DisplayName("New Password")]
        public InArgument<string> NewPassword { get; set; }

        public string strFilePath { get; set; }
        public string strPassword { get; set; }
        public string strNewPasswrod { get; set; }

        private ExcelObj._Application xlApp = null;
        private ExcelObj._Workbook xlWorkBook = null;
        private object misValue = System.Reflection.Missing.Value;

        protected override void Execute(CodeActivityContext context)
        {
            this.InitValues(context);

            this.doSetPassword();
        }

        private void InitValues(CodeActivityContext context)
        {
            this.strFilePath = this.FilePath.Get(context);
            this.strPassword = this.FilePassword.Get(context);
            this.strNewPasswrod = NewPassword.Get(context);
        }

        public void doSetPassword()
        {
            try
            {
                // New instantiation 
                xlApp = new ExcelObj.Application();

                if (!string.IsNullOrEmpty(strPassword))
                {
                    xlWorkBook = xlApp.Workbooks.Open(strFilePath, misValue, misValue, misValue, strPassword);
                }
                else
                {
                    xlWorkBook = xlApp.Workbooks.Open(strFilePath);
                }


                xlApp.DisplayAlerts = false;

                xlWorkBook.SaveAs(strFilePath, misValue, strNewPasswrod, misValue, misValue, misValue, ExcelObj.XlSaveAsAccessMode.xlNoChange,
    misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                this.ClearObject();

            }
            catch (Exception ex)
            {
                this.ClearObject();
                throw ex;

            }
        }

        protected void ClearObject()
        {
            if (xlWorkBook != null)
            {
                xlWorkBook.Close(true, misValue, misValue);

                base.releaseObject(xlWorkBook);
            }

            if (xlApp != null)
            {
                xlApp.Quit();

                base.releaseObject(xlApp);
            }

            base.ClearnGarbage();
        }
    }
}
*/