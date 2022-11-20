using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using Aspose.Words;
using System.IO;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;

public class Convert : MonoBehaviour
{
    public string _pathToDocx;
    public string _pathToPPTX;

    [ContextMenu("WordToPDF")]
    public void WordToPDF()   
    {
        var doc = new Document(_pathToDocx);
        doc.Save("C:\\Users\\Arvanito\\Documents\\DOCXPDF.pdf");
        print("Convert complete");
    }
    [ContextMenu("SlidesToPDF")]
    public void SlidesToPDF()
    {
        //Carrega a apresentação do power point via stream
        using (FileStream fileStreamInput = new FileStream(_pathToPPTX, FileMode.Open, FileAccess.ReadWrite))
        {
            //carrega o stream do ppt no Presentation
            using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
            {
                foreach (ISlide slide in pptxDoc.Slides)
                {
                    //Intera sobre as formas do powerpoint
                    foreach (IShape shape in slide.Shapes)
                    {
                        if (shape != null)
                        {
                            switch (shape.TextBody.Text)
                            {
                                case "Teste 1":
                                    shape.TextBody.Text = "Título";
                                    break;
                                case "Teste 2":
                                    shape.TextBody.Text = "Texto";
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }

                //Cria o PDF para fazer a transferencia
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    //Converte o power point para o pdf
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                    {
                        //Salva o PDF convertido em MemoryStream.
                        pdfDocument.Save(pdfStream);
                        pdfStream.Position = 0;
                    }

                    //Cria saida para o PDF
                    using (FileStream fileStreamOutput = File.Create("C:\\Users\\Arvanito\\Documents\\SLIDEPDF.pdf"))
                    {
                        //copia o pdf convertido para a saída
                        pdfStream.CopyTo(fileStreamOutput);
                        print("Convert complete");
                    }
                }
            }
        }
    }

    private void Update()
    {
        if(Input.GetKeyDown(KeyCode.W))
        {
            WordToPDF();
        }

        if (Input.GetKeyDown(KeyCode.P))
        {
            SlidesToPDF();
        }
    }
}
