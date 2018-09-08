namespace BalaReva.PowerPoint
{
    using Microsoft.Office.Core;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using PowerPointObj = Microsoft.Office.Interop.PowerPoint;
    using ReleaseObj = System.Runtime.InteropServices;

    public abstract class BasePowerPoint : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Select PPT File with path")]
        [DisplayName("File Path")]
        public InArgument<string> FilePath { get; set; }

        protected PowerPointObj._Application PptApplication = null;
        protected PowerPointObj._Presentation PptPresentation = null;

        protected CodeActivityContext Context { get; set; }

        protected string strFile { get; set; }
        protected bool IsAppOpened = false;
        protected bool IsPresentationOpened = false;

       protected int SlideCount
        {
            get
            {
                if (this.PptPresentation != null)
                {
                    return this.PptPresentation.Slides.Count;
                }
                else
                {
                    return 0;
                }
            }
        }


        protected void LoadValues(CodeActivityContext context)
        {
            this.Context = context;

            this.strFile = FilePath.Get(context);
        }

        protected void InitApplication()
        {
            this.PptApplication = new PowerPointObj.Application();
            this.IsAppOpened = true;
        }

        protected PowerPointObj.CustomLayout GetCustomLayout()
        {
            return PptPresentation.SlideMaster.CustomLayouts[PowerPointObj.PpSlideLayout.ppLayoutText];
        }


        protected void InitPresentation()
        {
            this.Validate();
            this.InitApplication();
            this.PptPresentation = PptApplication.Presentations.Open(strFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            this.IsPresentationOpened = true;
        }

        protected void SavePresentation()
        {
            if (this.IsPresentationOpened && this.PptPresentation != null)
            {
                this.PptPresentation.Save();
                this.PptPresentation.Close();
                this.IsPresentationOpened = false;
            }

            if (this.IsAppOpened && this.PptApplication != null)
            {
                this.PptApplication.Quit();
                this.IsAppOpened = false;
            }

            this.ClearObject();
        }

        protected void ClearObject()
        {

            if (this.PptPresentation != null)
            {
                if (this.IsPresentationOpened)
                {
                    this.PptPresentation.Close();
                    this.IsPresentationOpened = false;
                }

                ReleaseObj.Marshal.ReleaseComObject(PptPresentation);
                this.PptPresentation = null;
            }


            if (this.PptApplication != null)
            {
                if (this.IsAppOpened)
                {
                    this.PptApplication.Quit();
                    this.IsAppOpened = false;
                }

                ReleaseObj.Marshal.ReleaseComObject(PptApplication);
                this.PptApplication = null;
            }

            this.ClearnGarbage();
        }


        private void Validate()
        {
            if (!System.IO.File.Exists(strFile))
            {
                throw new Exception("File is not exists");
            }
        }

        private void ClearnGarbage()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
