using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BaseOffice
{
    public abstract class RevitCommandBase : IExternalCommand
    {
        protected RevitContext Context { get; private set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Cria um novo contexto atualizado sempre que o comando for executado
            Context = new RevitContext(commandData.Application);

            // Chama a implementação específica do comando herdado
            return ExecuteCommand(ref message, elements);
        }

        // Método abstrato que os comandos devem implementar
        protected abstract Result ExecuteCommand(ref string message, ElementSet elements);
    }

}
