using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.ApplicationServices;

namespace UniBimTablas
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]

    public class SincronizarId : IExternalCommand
    {
        public static Document _doc;
        public static Application _app;
        public static bool _itemizar;
        public static bool _numerar;
        public static string _prefijo;
        public static string _sufijo;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            TaskDialog cuadro = new TaskDialog("Sincronizar Ids");//, "Esta herramienta permite rellenar un parámetro de ejemplar llamado 'id' con el valor del Id del elemento");
            cuadro.MainInstruction = "Sincronizar Ids de ejemplares de familia";
            cuadro.MainContent = "Esta herramienta permite rellenar un parámetro de ejemplar llamado 'id' con el valor del Id del elemento. Nota: el parámetro debe ser de tipo ENTERO ¿Desea continuar?";
            cuadro.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Continuar");
            cuadro.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Cancelar");
            //cuadro.CommonButtons = TaskDialogCommonButtons.Cancel;
            //cuadro.DefaultButton = TaskDialogResult.Cancel;
            TaskDialogResult resultado = cuadro.Show();
            if (TaskDialogResult.CommandLink1 == resultado)
            {
                List<Element> instancias = Tools.GetInstances(doc);
                Tools.SincronizarIdsEjemplares(doc, instancias);
                TaskDialog.Show("Sincronizar Ids", "Proceso terminado con éxito");
            }

            return Result.Succeeded;
        }
    }
}
