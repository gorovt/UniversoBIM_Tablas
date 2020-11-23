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

    public class CrearTablas : IExternalCommand
    {
        public static Document _doc;
        public static Application _app;
        public static bool _itemizar;
        public static bool _linked;
        public static bool _numerar;
        public static string _prefijo;
        public static string _sufijo;
        public static bool _marca;
        public static bool _comentarios;
        public static bool _descripcion;
        public static bool _modelo;
        public static bool _marcaTipo;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            _itemizar = false;
            _linked = false;
            _numerar = true;
            _prefijo = "-";
            _sufijo = "";
            _marca = false;
            _comentarios = false;
            _descripcion = true;
            _modelo = false;
            _marcaTipo = false;

            System.Windows.Forms.DialogResult resultado = (new frmCrearTablas(uidoc.Document)).ShowDialog();

            if (resultado == System.Windows.Forms.DialogResult.OK)
            {
                List<ViewSchedule> tablas = new List<ViewSchedule>();

                // Creamos y abrimos una transacción
                using (Transaction t = new Transaction(doc, "Crear Tablas"))
                {
                    t.Start();
                    // Ordenar Categorias
                    Tools.categorias = Tools.categorias.OrderBy(x => x.Name).ToList();

                    int count = 1;
                    int countMats = 1;
                    try
                    {
                        // Crear una Tabla para cada Categoría
                        foreach (Category cat in Tools._selectedCategories)
                        {
                            string nombre = "";
                            if (_numerar)
                            {
                                nombre += count.ToString();
                            }
                            if (_prefijo != "")
                            {
                                nombre += _prefijo;
                            }
                            nombre += " " + cat.Name;
                            if (_sufijo != "")
                            {
                                nombre += _sufijo;
                            }

                            ViewSchedule tabla = ViewSchedule.CreateSchedule(doc, cat.Id);
                            tabla.Name = nombre;
                            count++;
                            // ver 1.4 MEP
                            bool mep = false;
                            bool pipeSystem = false;
                            bool ductSystem = false;
                            if (cat.Id.IntegerValue == -2008044) //Pipes
                            {
                                mep = true;
                                pipeSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008049) //PipeFittings
                            {
                                mep = true;
                                pipeSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008055) //Pipe Accesories
                            {
                                mep = true;
                                pipeSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008050) //FlexPipe
                            {
                                //mep = true; // DEBE SER TAMAÑO TOTAL
                                pipeSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008000) // Duct
                            {
                                mep = true;
                                ductSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008010) // DuctFitting
                            {
                                mep = true;
                                ductSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008016) // Duct Accesories
                            {
                                mep = true;
                                ductSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008020) // FlexDuct
                            {
                                //mep = true; // DEBE SER TAMAÑO TOTAL
                                ductSystem = true;
                            }
                            if (cat.Id.IntegerValue == -2008132) // Conduit
                            {
                                mep = true;
                            }
                            if (cat.Id.IntegerValue == -2008128) // ConduitFitting
                            {
                                mep = true;
                            }
                            if (cat.Id.IntegerValue == -2008130) // CableTray
                            {
                                mep = true;
                            }
                            if (cat.Id.IntegerValue == -2008126) // CableTrayFittings
                            {
                                mep = true;
                            }
                            Tools.FillSchedule(tabla, _itemizar, _linked, mep, pipeSystem, ductSystem);
                            // Phase Filter
                            Parameter tablaFilter = tabla.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER);
                            tablaFilter.Set(new ElementId(-1));
                            tablas.Add(tabla);
                        }
                        // Create Material Schedules
                        foreach (Category cat in Tools._selectedMatCategories)
                        {
                            int createMat = 0;

                            if (cat.HasMaterialQuantities)// && cat.Material != null)
                            {
                                if (cat.Id == new ElementId(BuiltInCategory.OST_Walls))
                                {
                                    createMat = 1;
                                }
                                if (cat.Id == new ElementId(BuiltInCategory.OST_Floors))
                                {
                                    createMat = 1;
                                }
                                if (cat.Id == new ElementId(BuiltInCategory.OST_Roofs))
                                {
                                    createMat = 1;
                                }
                                if (cat.Id == new ElementId(BuiltInCategory.OST_Ceilings))
                                {
                                    createMat = 1;
                                }
                                if (cat.Id == new ElementId(BuiltInCategory.OST_BuildingPad))
                                {
                                    createMat = 1;
                                }

                                if (createMat == 1)
                                {
                                    string nombre = "M";
                                    if (_numerar)
                                    {
                                        nombre += countMats.ToString();
                                    }
                                    if (_prefijo != "")
                                    {
                                        nombre += _prefijo;
                                    }
                                    nombre += " " + cat.Name;
                                    if (_sufijo != "")
                                    {
                                        nombre += _sufijo;
                                    }

                                    ViewSchedule tablaMat = ViewSchedule.CreateMaterialTakeoff(doc, cat.Id);
                                    tablaMat.Name = nombre + " (Materiales)";
                                    // Phase Filter
                                    Parameter tablaFilter = tablaMat.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER);
                                    tablaFilter.Set(new ElementId(-1));
                                    countMats++;
                                    Tools.FillScheduleMaterial(tablaMat, _itemizar, _linked);
                                    tablas.Add(tablaMat);
                                }
                            }
                        }
                        t.Commit();
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Universo BIM", "Error: " + ex.Message);
                        t.RollBack();
                    }
                    
                }
                TaskDialog.Show("Universo BIM", "Se han creado " + tablas.Count + " Tablas");
            }

            return Result.Succeeded;
        }
    }
}
