using System.Reflection;
using word = Microsoft.Office.Interop.Word;
using PruebaComceptoCrearDocumentos.Modelos;
using Microsoft.Office.Interop.Word;

namespace PruebaComceptoCrearDocumentos.Data
{
    public class GenerarDocumento
    {

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object toFindText, object replaceWithText)
        {
            object matchCase = true;

            object matchwholeWord = true;

            object matchwildCards = false;

            object matchSoundLike = false;

            object nmatchAllforms = false;

            object forward = true;

            object format = false;

            object matchKashida = false;

            object matchDiactitics = false;

            object matchAlefHamza = false;

            object matchControl = false;

            object read_only = false;

            object visible = true;

            object replace = -2;

            object wrap = 1;

            wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                            ref matchwholeWord, ref matchwildCards, ref matchSoundLike,

                                            ref nmatchAllforms, ref forward,

                                            ref wrap, ref format, ref replaceWithText,

                                                ref replace, ref matchKashida,

                                            ref matchDiactitics, ref matchAlefHamza,

                                             ref matchControl);

        }
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


                this.cambiarContenido(myWordDoc, "NameStuden", DatosDocumento.NameAlumno);
                this.cambiarContenido(myWordDoc, "TextDatatime", DatosDocumento.Fecha);
                this.cambiarContenido(myWordDoc, "nameDirector", DatosDocumento.NameDirector);
                this.cambiarContenido(myWordDoc, "NameTutor", DatosDocumento.NameTutor);

                myWordDoc.ExportAsFixedFormat(SaveAs.ToString(), FormatoExportado);

                
                myWordDoc.Close(null,null,null);
                wordApp.Quit();
            }
        }

        private void cambiarContenido(Microsoft.Office.Interop.Word.Document wordApp, string tituloContecontrol, string nuevoContenido)
        {




            ContentControls objcc = wordApp.SelectContentControlsByTitle(tituloContecontrol);
            
            int contar = objcc.Count;

            if (contar > 1)
            {
                for (int i = 1; i <= contar; i++)
                {
                    objcc[i].Title = nuevoContenido;
                    objcc[i].Range.Text = nuevoContenido;
                }
            }
            else {

                objcc[contar].Title = nuevoContenido;
                objcc[contar].Range.Text = nuevoContenido;
            }
               
                
            

            
        }


    }
}
