using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseOffice.Startup
{
    internal static class UIBuilder
    {
        internal static void BuildUI(UIControlledApplication application)
        {
            RibbonPanel panelMain = RibbonManager.CriarRibbonPanel(application, "Main");

            RibbonManager.CriarPushButton
            ( "Dev", "Dev",
            "BaseOffice.Dev",
            panelMain,
            "Modo Desenvolvedor",
            "dev.ico");


            //RibbonManager.CriarPushButton
            //("NomeInterno", "NomeExibido",
            //"NameSpace.Classe",
            //RibbonPanel,
            //"Dica de uso",
            //"imagem.ico");
        }
    }
}
