namespace BalaReva.Excel
{
    using System;
    using System.Activities;
    using System.ComponentModel;
    using ExcelObj = Microsoft.Office.Interop.Excel;
    using ReleaseObj = System.Runtime.InteropServices;

    public abstract class BaseExcelNew : CodeActivity
    {
        [Category("Input"), RequiredArgument]
        [Description("File path")]
        [DisplayName("File Path")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [Description("File password")]
        [DisplayName("File Password")]
        public InArgument<string> FilePassword { get; set; }

        [Category("Input"), RequiredArgument]
        [Description("Sheet Name")]
        [DisplayName("Sheet Name")]
        public InArgument<string> SheetName { get; set; }

        protected string strFilePath { get; set; }
        protected string strSheetName { get; set; }
        protected string strPassword { get; set; }

        protected ExcelObj._Application xlApp = null;
        protected ExcelObj._Workbook xlWorkBook = null;
        protected ExcelObj._Worksheet xlWorkSheet = null;
        protected object misValue = System.Reflection.Missing.Value;
        protected Int16 SheetIndex { get; set; }

        protected Int32 SheetCount
        {
            get
            {
                return this.xlWorkBook.Sheets.Count;
            }
        }


        protected bool IsWorkBookOpened = false;
        private bool IsAppOpened = false;
        protected bool IsWorkSheetOpened = false;

        protected void LoadVariables(CodeActivityContext context)
        {
            this.strFilePath = this.FilePath.Get(context);
            this.strSheetName = this.SheetName.Get(context);
            this.strPassword = this.FilePassword.Get(context);
        }

        protected void InitApp()
        {
            // New instantiation 
            xlApp = new ExcelObj.Application();
            xlApp.DisplayAlerts = false;
            this.IsAppOpened = true;
        }


        protected void InitWorkBook()
        {
            // New instantiation 
            this.InitApp();

            if (!string.IsNullOrEmpty(strPassword))
            {
                xlWorkBook = xlApp.Workbooks.Open(strFilePath, misValue, misValue, misValue, strPassword);
            }
            else
            {
                xlWorkBook = xlApp.Workbooks.Open(strFilePath);
            }

            this.IsWorkBookOpened = true;
        }

        protected void InitWorkSheet()
        {
            IsWorkSheetOpened = false;

            SheetIndex = 0;

            this.InitWorkBook();

            // to get the sheet index
            for (Int16 i = 1; i <= xlWorkBook.Worksheets.Count; i++)
            {
                if (xlWorkBook.Worksheets[i].Name == this.strSheetName)
                {
                    SheetIndex = i;
                    break;
                }
            }

            if (this.SheetIndex > 0)
            {
                // initialize the work sheet object
                xlWorkSheet = (ExcelObj.Worksheet)xlWorkBook.Worksheets.get_Item(this.SheetIndex);
                IsWorkSheetOpened = true;
            }
            else
            {
                this.ClearObject();

                throw new Exception("Invalid Sheet Name");
            }
        }


        protected ExcelObj.Range GetSheetRange(string strRange)
        {
            ExcelObj.Range xlRange = null;

            try
            {
                if (this.IsWorkSheetOpened)
                {
                    xlRange = this.xlWorkSheet.Range[strRange];
                }
            }
            catch (Exception)
            {
                xlRange = null;
            }

            return xlRange;
        }

        protected void SaveWorkBook(bool blnSave)
        {
            // save and close the workbook
            if (this.xlWorkBook != null)
            {
                if (blnSave && this.IsWorkBookOpened)
                {
                    xlWorkBook.Save();
                }

                xlWorkBook.Close(true, misValue, misValue);
                this.IsWorkBookOpened = false;
                this.IsWorkSheetOpened = false;
            }

            if (xlApp != null && this.IsAppOpened)
            {
                xlApp.Quit();
                this.IsAppOpened = false;
            }

            // close the object
            this.ClearObject();
        }

        protected void ClearObject()
        {
            if (this.xlWorkSheet != null)
            {
                ReleaseObj.Marshal.ReleaseComObject(xlWorkSheet);
                this.IsWorkSheetOpened = false;
                this.xlWorkSheet = null;
            }

            if (this.xlWorkBook != null)
            {
                if (this.IsWorkBookOpened)
                {
                    xlWorkBook.Close(false, misValue, misValue);
                    this.IsWorkBookOpened = false;
                }

                ReleaseObj.Marshal.ReleaseComObject(xlWorkBook);
                this.xlWorkBook = null;
            }

            if (this.xlApp != null)
            {
                if (this.IsAppOpened)
                {
                    xlApp.Quit();
                    this.IsAppOpened = false;
                }

                ReleaseObj.Marshal.ReleaseComObject(xlApp);
                this.xlApp = null;
            }

            this.ClearnGarbage();
        }

        //cleanup the Garbage.
        private void ClearnGarbage()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
