using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System.Windows.Media.Imaging;

namespace UniBimTablas
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]

    public class App : IExternalApplication
    {
        static Autodesk.Revit.DB.AddInId m_appId = new Autodesk.Revit.DB.AddInId(new Guid("ef0991ad-d7e7-4a33-b45c-050c3b5e1970"));
        //get the absolute path of this assembly
        static string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            AddMenu(application);
            return Autodesk.Revit.UI.Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        private void AddMenu(UIControlledApplication app)
        {
            Autodesk.Revit.UI.RibbonPanel pannel = app.CreateRibbonPanel("UniBIM Tablas 1.50");
            //IMPORTANT NOTE: las imagenes de los botones deben tener su propiedad "Build Action" como Resource.
            string pathResource = "pack://application:,,,/UniBimTablas;component/Resources/";

            // Botones
            PushButtonData button1 = new PushButtonData("ubTablas01", "Exportar Tablas",
                ExecutingAssemblyPath, "UniBimTablas.ExportarTablas");
            button1.LargeImage = new BitmapImage(new Uri(pathResource + "Excel32.png"));

            PushButtonData button2 = new PushButtonData("ubTablas02", "Crear Tablas",
                ExecutingAssemblyPath, "UniBimTablas.CrearTablas");
            button2.LargeImage = new BitmapImage(new Uri(pathResource + "table32.png"));

            PushButtonData button3 = new PushButtonData("dbTablas03", "Sincronizar Ids",
                ExecutingAssemblyPath, "UniBimTablas.SincronizarId");
            button3.LargeImage = new BitmapImage(new Uri(pathResource + "tag.png"));

            // PullDown Buttons
            pannel.AddItem(button2);
            pannel.AddItem(button1);
            pannel.AddItem(button3);
        }
    }
}
