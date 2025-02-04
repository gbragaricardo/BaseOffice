using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;


namespace BaseOffice
{
    [Transaction(TransactionMode.Manual)]
    internal class Dev : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RevitContext context = new RevitContext(commandData.Application);

            using (Transaction transacao = new Transaction(context.Doc, "DEV"))
            {

                transacao.Start();

                var tab = context.UiApp.GetRibbonPanels("BaseOffice");

                foreach (var panel in tab)
                {
                     var botoes = panel.GetItems();

                    foreach (var item in botoes)
                    {

                        item.Enabled = true;
                    }
                }


                TaskDialog.Show("DEV", "Botoes Ativados");

                transacao.Commit();

                return Result.Succeeded;
            }
        }
    }
}