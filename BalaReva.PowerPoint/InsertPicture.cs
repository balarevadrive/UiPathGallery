namespace BalaReva.PowerPoint
{
    using Design;
    using Microsoft.Office.Core;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("Insert Picture")]
    [Designer(typeof(InsertPictureDesign))]
    public class InsertPicture : BasePowerPoint
    {
        [Category("Input"), RequiredArgument]
        [Description("Select Image to insert")]
        [DisplayName("Image Path")]
        public InArgument<string> ImagePath { get; set; }

        [Category("Input"), RequiredArgument]
        [Description("Slide Index starts with one")]
        [DisplayName("Slide Index")]
        public InArgument<int> SlideIndex { get; set; } = 1;

        [Category("Input")]
        [RequiredArgument]
        public ImageSize PictureSize { get; set; } = new ImageSize();

        private int intSlideIndex { get; set; }
        private string strImage { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                base.LoadValues(context);

                this.intSlideIndex = SlideIndex.Get(context);
                this.strImage = ImagePath.Get(context);

                this.DoInsertPicture();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoInsertPicture()
        {
            this.ValidateInserPicture();

            try
            {
                base.InitPresentation();

                if (base.PptPresentation != null)
                {
                    if (base.SlideCount >= intSlideIndex)
                    {
                        base.PptPresentation.Slides[intSlideIndex].Shapes.AddPicture
                        (this.strImage, MsoTriState.msoFalse, MsoTriState.msoTrue,
                        PictureSize.Left, PictureSize.Top, PictureSize.Width, PictureSize.Height);

                        this.SavePresentation();
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

        private void ValidateInserPicture()
        {
            if (!System.IO.File.Exists(strImage))
            {
                throw new Exception("Invalid Image Path");
            }
        }
    }
}
