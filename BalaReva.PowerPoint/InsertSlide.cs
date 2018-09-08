namespace BalaReva.PowerPoint
{
    using BalaReva.PowerPoint.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;
    using PowerPointObj = Microsoft.Office.Interop.PowerPoint;

    [DisplayName("Insert Slide")]
    [Designer(typeof(InsertSlideDesign))]
    public class InsertSlide : BasePowerPoint
    {
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadValues(context);

                this.DoInsertSlide();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoInsertSlide()
        {
            try
            {
                base.InitPresentation();

                if (this.PptPresentation != null)
                {
                    PowerPointObj.CustomLayout customLayout = base.GetCustomLayout();

                    this.PptPresentation.Slides.AddSlide(base.SlideCount + 1, customLayout);

                    base.SavePresentation();
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
