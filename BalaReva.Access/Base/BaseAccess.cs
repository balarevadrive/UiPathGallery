using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AccessObj = Microsoft.Office.Interop.Access;
using ReleaseObj = System.Runtime.InteropServices;

namespace BalaReva.Access
{
    public abstract class BaseAccess : CodeActivity
    {

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter the File Path")]
        public InArgument<string> FilePath { get; set; }


        [Category("Input")]
        [Description("Enter the FilePassword")]
        public InArgument<string> Password { get; set; }

        protected CodeActivityContext Context { get; set; }

        protected AccessObj.Application AccApp = null;
        protected bool IsAppOpened { get; set; } = false;

        private string StrFilePath;
        private string StrPassword;

        protected void LoadVariables(CodeActivityContext context)
        {
            this.Context = context;

            this.StrFilePath = this.FilePath.Get(context);
            this.StrPassword = this.Password.Get(context);
        }

        protected void InitApp()
        {
            this.AccApp = new AccessObj.Application();
            this.IsAppOpened = true;

            if (string.IsNullOrEmpty(StrPassword))
            {
                AccApp.OpenCurrentDatabase(StrFilePath, false);
            }
            else
            {
                AccApp.OpenCurrentDatabase(StrFilePath, false, StrPassword);
            }
        }

        protected void Save()
        {
            if (AccApp != null && this.IsAppOpened)
            {
                AccApp.Quit();
                this.IsAppOpened = false;
            }

            this.ClearObject();
        }

        protected void ClearObject()
        {
            if (this.AccApp != null)
            {
                if (this.IsAppOpened)
                {
                    AccApp.Quit();
                    this.IsAppOpened = false;
                }

                ReleaseObj.Marshal.ReleaseComObject(AccApp);
                this.AccApp = null;
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
