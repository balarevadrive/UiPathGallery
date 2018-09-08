// https://www.nuget.org/packages/sautinsoft.document/

namespace BalaReva.PDF
{
    using SautinSoft.Document;
    using System;
    using System.Activities;
    using System.ComponentModel;

    [DisplayName("PDF To Word")]
    public class PdfToWord : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("PDF File Path")]
        public InArgument<string> PDFFilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Word File Path")]
        public InArgument<string> WordFilePath { get; set; }

        [Category("Input")]
        [Description("Word File password")]
        public InArgument<string> WordFilePassword { get; set; }

        string PDFPath { get; set; }
        string WordPath { get; set; }
        string WordPwd { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            PDFPath = PDFFilePath.Get(context);
            WordPath = WordFilePath.Get(context);
            WordPwd = WordFilePassword.Get(context);

            this.ConvertPdfToWord(PDFPath, WordPath, WordPwd);
        }

        private void ConvertPdfToWord(string pdfFile, string docxFile, string docxPass)
        {
            try
            {
                PdfLoadOptions pdfOptions = new PdfLoadOptions();
                pdfOptions.ConversionMode = PdfConversionMode.Flowing;
                pdfOptions.DetectTables = true;
                pdfOptions.RasterizeVectorGraphics = true;
                if (!string.IsNullOrEmpty(docxPass))
                {
                    pdfOptions.Password = docxPass;
                }

                DocumentCore pdf = DocumentCore.Load(pdfFile, pdfOptions);
                pdf.Save(docxFile);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

    }
}
