namespace BalaReva.Excel.Sheets
{
    using Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using WorkBook;
    using ExcelObj = Microsoft.Office.Interop.Excel;
    using ReleaseObj = System.Runtime.InteropServices;

    [DisplayName("Copy To File")]
    [Designer(typeof(Excel2Selection))]
    public class CopyToFile : BaseExcelNew
    {
        [Category("Input"), RequiredArgument]
        [Description("New File path")]
        [DisplayName("New File Path")]
        public InArgument<string> NewFilePath { get; set; }

        [Category("Input")]
        [Description("New File Password")]
        [DisplayName("New File Password")]
        public InArgument<string> NewFilePassword { get; set; }


        [Category("Input")]
        [Description("New Sheet Name")]
        [DisplayName("New Sheet Name")]
        public InArgument<string> NewSheetName { get; set; }

        [Category("Input")]
        [Description("New file is not exists,it creates automatically")]
        [DisplayName("Auto File Creation"), DefaultValue(false)]
        public InArgument<bool> AutoFileCreation { get; set; } = false;

        private string StrNewSheetName { get; set; }
        private string StrNewFilePath { get; set; }
        private string StrNewFilePassword { get; set; }
        private bool IsAutoCreate { get; set; }

        private ExcelObj.Workbook xlWorkBookTarget = null;
        private ExcelObj.Worksheet xlWorkSheetTarget = null;
        private bool IsWorkBookTargetOpened = false;


        protected override void Execute(CodeActivityContext context)
        {
            base.LoadVariables(context);

            StrNewSheetName = NewSheetName.Get(context);
            StrNewFilePath = NewFilePath.Get(context);
            StrNewFilePassword = NewFilePassword.Get(context);
            IsAutoCreate=AutoFileCreation.Get(context);

            this.DoCopyToFile();
        }

        private void DoCopyToFile()
        {
            try
            {
                this.FileValidation();

                base.InitWorkSheet();

                if (base.xlWorkSheet != null)
                {
                    this.AddSheetToExcel(base.xlWorkSheet);

                    base.ClearObject();
                }
            }
            catch (Exception ex)
            {
                this.ClearTargetObject();
                base.ClearObject();

                throw ex;
            }
        }

        private void AddSheetToExcel(ExcelObj._Worksheet NewSheet)
        {

            if (!string.IsNullOrEmpty(this.StrNewFilePassword))
            {
                this.xlWorkBookTarget = base.xlApp.Workbooks.Open(StrNewFilePath, misValue, misValue, misValue, StrNewFilePassword);
                this.IsWorkBookTargetOpened = true;
            }
            else
            {
                this.xlWorkBookTarget = base.xlApp.Workbooks.Open(StrNewFilePath);
                this.IsWorkBookTargetOpened = true;
            }

            this.xlWorkSheetTarget = (ExcelObj.Worksheet)this.xlWorkBookTarget.Sheets[this.xlWorkBookTarget.Sheets.Count];


            NewSheet.Copy(Type.Missing, xlWorkSheetTarget);

            if (StrNewSheetName != null && !string.IsNullOrEmpty(StrNewSheetName))
            {
                this.xlWorkSheetTarget = (ExcelObj.Worksheet)this.xlWorkBookTarget.Sheets.Item[this.xlWorkBookTarget.Sheets.Count];
                this.xlWorkSheetTarget.Name = StrNewSheetName;
            }

            xlWorkBookTarget.Save();
            xlWorkBookTarget.Close();
            this.IsWorkBookTargetOpened = false;

            ReleaseObj.Marshal.ReleaseComObject(NewSheet);
            NewSheet = null;

            this.ClearTargetObject();
        }

        protected void ClearTargetObject()
        {
            if (this.xlWorkSheetTarget != null)
            {
                ReleaseObj.Marshal.ReleaseComObject(xlWorkSheetTarget);
                this.xlWorkSheetTarget = null;
            }

            if (this.xlWorkBookTarget != null)
            {
                if (this.IsWorkBookTargetOpened)
                {
                    xlWorkBookTarget.Close(false, misValue, misValue);
                    this.IsWorkBookTargetOpened = false;
                }

                ReleaseObj.Marshal.ReleaseComObject(xlWorkBookTarget);
                this.xlWorkBookTarget = null;
            }
        }

        private void FileValidation()
        {
            if (!System.IO.File.Exists(StrNewFilePath))
            {
                if (IsAutoCreate)
                {
                    CreateWorkBook createWorkBook = new CreateWorkBook();
                    createWorkBook.FilePath = StrNewFilePath;
                    createWorkBook.FilePassword = StrNewFilePassword;
                    createWorkBook.SheetName = "Sheet1";

                    WorkflowInvoker.Invoke(createWorkBook);
                }
                else
                {
                    throw new Exception("File doesn't exists");
                }
            }
        }
    }
}
