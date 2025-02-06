using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using BaseOffice.Startup;

namespace BaseOffice
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]

    public class Application : IExternalApplication
    {

        public Result OnShutdown(UIControlledApplication application)
        {
            EventoModificacao.Finalizar(application); // Registra o evento de modificação
            return Result.Succeeded;
        }


        public Result OnStartup(UIControlledApplication application)
        {
            EventoModificacao.Inicializar(application); // Registra o evento de modificação

            try {AddinAppLoader.StartupMain(application);}
            catch{TaskDialog.Show("ProjetaHDR", "Erro ao inicializar Plugin ProjetaHDR");}

            return Result.Succeeded;
        }
    }
}
