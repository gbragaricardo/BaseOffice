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
    internal class Template : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RevitContext context = new RevitContext(commandData.Application);

            using (Transaction transacao = new Transaction(context.Doc, "DEV"))
            {

                transacao.Start();

                TaskDialog.Show("DEV", "Hello World");

                transacao.Commit();

                return Result.Succeeded;
            }
        }
    }
}