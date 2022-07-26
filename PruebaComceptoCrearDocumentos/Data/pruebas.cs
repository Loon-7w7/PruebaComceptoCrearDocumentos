using Microsoft.Office.Interop.Word;
using PruebaComceptoCrearDocumentos.Modelos;
using System.Reflection;
using word = Microsoft.Office.Interop.Word;

namespace PruebaComceptoCrearDocumentos.Data
{
    public class pruebas
    {
        private readonly WdContentControlType wdContentControlText;

        public void CreateWordDocument(object filename, object SaveAs, DataDocumente DatosDocumento)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            Microsoft.Office.Interop.Word.WdExportFormat FormatoExportado = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;

            Microsoft.Office.Interop.Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;

                object isvisible = false;

                wordApp.Visible = false;
                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                     ref missing, ref missing, ref missing, ref missing);

                myWordDoc.Activate();

                cambiarContenido(myWordDoc);


                myWordDoc.ExportAsFixedFormat(SaveAs.ToString(), FormatoExportado);


                myWordDoc.Close(null,null,null);
                wordApp.Quit();
            }
        }

        private void cambiarContenido(Microsoft.Office.Interop.Word.Document wordApp) {



            ContentControls objcc = wordApp.SelectContentControlsByTitle("Titulo");
 
                objcc[1].Title = "Cambiado ";
                objcc[1].Range.Text = "cambia puto";
            


        }
    }
}
