using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Events;

namespace BaseOffice
{
    public class EventoModificacao
    {
        public static void Inicializar(UIControlledApplication application)
        {
            application.ControlledApplication.DocumentChanged += OnDocumentChanged;
        }

        private static void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            
        }

        public static void Finalizar(UIControlledApplication application)
        {
            application.ControlledApplication.DocumentChanged -= OnDocumentChanged;
        }
    }

}