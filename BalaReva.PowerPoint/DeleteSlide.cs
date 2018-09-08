namespace BalaReva.PowerPoint
{
    using BalaReva.PowerPoint.Design;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Delete Slide")]
    [Designer(typeof(InsertSlideDesign))]
    public class DeleteSlide : BasePowerPoint
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Index start wiht one")]
        [DisplayName("Slide Index")]
        public InArgument<int> SlideIndex { get; set; } = 1;

        private int intSlideIndex { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadValues(context);

                this.intSlideIndex = SlideIndex.Get(context);

                this.DoDeleteSlide();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoDeleteSlide()
        {
            try
            {
                base.InitPresentation();

                if (base.PptPresentation != null)
                {
                    if (base.SlideCount >= intSlideIndex)
                    {
                        base.PptPresentation.Slides[intSlideIndex].Delete();

                        base.SavePresentation();
                    }
                    else
                    {
                        base.ClearObject();
                        throw new Exception("Invalid slide index");
                    }
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
