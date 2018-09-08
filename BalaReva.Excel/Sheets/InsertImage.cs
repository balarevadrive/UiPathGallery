/*namespace BalaReva.Excel.Sheets
{
    using Microsoft.Office.Core;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;

    public class InsertImage : BaseExcelNew
    {
        [Category("Input")]
        [Description("Image full Path")]
        [DisplayName("Image Path")]
        public InArgument<string> ImagePath { get; set; }

        private string strImagePath { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            strImagePath = ImagePath.Get(context);

            this.DoProtect();
        }

        private void DoProtect()
        {
            try
            {
                base.InitWorkSheet();

                if (base.xlWorkSheet != null)
                {
                    ExcelObj.Worksheet worksheet = (ExcelObj.Worksheet)base.xlWorkSheet;
                    //this.xlWorkSheet.Shapes.AddPicture(strImagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 23, 34, 23, 23);

                    worksheet.Shapes.AddPicture(strImagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 24, 234, 234, 242);

                    base.SaveWorkBook(true);
                }
            }
            catch (Exception ex)
            {
                base.ClearObject();
                throw ex;
            }
        }

        private void ValidateImageFile()
        {
            if (!System.IO.File.Exists(strImagePath))
            {
                throw new Exception("Invalid Image path");
            }
        }
    }
}*/
