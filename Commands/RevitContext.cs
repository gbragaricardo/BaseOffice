using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BaseOffice
{
    public class RevitContext
    {
        public UIApplication UiApp { get; private set; }
        public UIDocument UiDoc { get; private set; }
        public Autodesk.Revit.ApplicationServices.Application App { get; private set; }
        public Document Doc { get; private set; }

        public RevitContext(UIApplication uiApp)
        {
            UiApp = uiApp;
            UiDoc = uiApp.ActiveUIDocument;
            App = uiApp.Application;
            Doc = UiDoc.Document;
        }
    }
}