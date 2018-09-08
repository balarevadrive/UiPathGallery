namespace BalaReva.PowerPoint
{
    using System.ComponentModel;

    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ImageSize : Interfaces.IImageSize
    {
        public ImageSize()
        {
            this.Left = 50;
            this.Top = 10;
            this.Width = 500;
            this.Height = 300;
        }

        public float Height { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
    }
}
