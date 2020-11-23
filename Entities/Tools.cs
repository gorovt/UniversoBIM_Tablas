using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.ApplicationServices;
using System.Windows.Forms;
//using Novacode;
using ClosedXML.Excel;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace UniBimTablas
{
    public class Tools
    {
        public static string version = "ver 1.5";
        public static List<Category> categorias = new List<Category>();
        public static List<Category> _selectedCategories = new List<Category>();
        public static List<Category> _selectedMatCategories = new List<Category>();

        #region GET AND COLLECTORS
        public static List<Element> GetInstances(Document doc)
        {
            List<Element> lstElement = new List<Element>();
            FilteredElementCollector col = new FilteredElementCollector(doc);
            List<Element> elementos = col.WhereElementIsNotElementType().ToList();
            lstElement = (from elem in elementos
                          where elem.Category != null
                          && elem.CreatedPhaseId.IntegerValue != -1
                          //&& elem.Category.Id != new ElementId(BuiltInCategory.OST_StairsRailing)
                          select elem).ToList();

            return lstElement;
        }
        public static List<Element> getFamilies(Document _doc)
        {
            List<Element> lstFamilies = new List<Element>();
            FilteredElementCollector col = new FilteredElementCollector(_doc);
            List<Element> lstElement = col.OfClass(typeof(FamilySymbol)).ToList();
            foreach (Element elem in lstElement)
            {
                if (elem.Category.CategoryType == CategoryType.Model && !lstFamilies.Exists(x => x.Name == elem.Name))
                    lstFamilies.Add(elem);
            }
            return lstFamilies;
        }
        public static List<Level> getLevels(Document _doc)
        {
            List<Level> lstLevels = new List<Level>();
            FilteredElementCollector col = new FilteredElementCollector(_doc);
            List<Element> lst = col.OfClass(typeof(Level)).ToList();
            foreach (Element elem in lst)
            {
                Level lev = elem as Level;
                lstLevels.Add(lev);
            }
            return lstLevels;
        }
        public static List<Category> getCategories(Document _doc)
        {
            List<Category> lstCategories = new List<Category>();
            foreach (Element elem in getFamilies(_doc))
            {
                FamilySymbol fam = elem as FamilySymbol;
                if (fam != null && !lstCategories.Exists(x => x.Name == fam.Category.Name))
                {
                    lstCategories.Add(fam.Category);
                }
            }
            lstCategories = lstCategories.OrderBy(x => x.Name).ToList();
            return lstCategories;
        }
        public static List<PipingSystemType> getSistemasTuberias(Document doc)
        {
            List<PipingSystemType> lstSistemas = new List<PipingSystemType>();
            FilteredElementCollector col = new FilteredElementCollector(doc);
            List<Element> lst = col.OfClass(typeof(PipingSystemType)).ToList();
            foreach (Element elem in lst)
            {
                PipingSystemType sistemType = elem as PipingSystemType;
                lstSistemas.Add(sistemType);
            }
            return lstSistemas;
        }
        public static List<PipeType> getTiposTuberia(Document doc)
        {
            List<PipeType> lstTipos = new List<PipeType>();
            FilteredElementCollector col = new FilteredElementCollector(doc);
            List<Element> lst = col.OfClass(typeof(PipeType)).ToList();
            foreach (Element elem in lst)
            {
                PipeType type = elem as PipeType;
                lstTipos.Add(type);
            }
            return lstTipos;
        }
        public static PipingSystemType getSystemByName(Document doc, string name)
        {
            PipingSystemType system = null;
            foreach (PipingSystemType type in getSistemasTuberias(doc))
            {
                if (type.Name == name)
                {
                    system = type;
                }
            }
            return system;
        }
        public static PipeType getPipeTypeByName (Document doc, string name)
        {
            PipeType type = null;
            foreach (PipeType pipe in getTiposTuberia(doc))
            {
                if (pipe.Name == name)
                {
                    type = pipe;
                }
            }
            return type;
        }
        public static Level getLevelByName(Document doc, string name)
        {
            Level level = null;
            foreach (Level lvl in getLevels(doc))
            {
                if (lvl.Name == name)
                {
                    level = lvl;
                }
            }
            return level;
        }
        public static List<ViewSchedule> getViewSchedules(Document doc)
        {
            List<ViewSchedule> lstTablas = new List<ViewSchedule>();
            FilteredElementCollector col = new FilteredElementCollector(doc);
            List<Element> lstElem = col.OfClass(typeof(ViewSchedule)).ToList();
            foreach (Element elem in lstElem)
            {
                ViewSchedule tabla = elem as ViewSchedule;
                lstTablas.Add(tabla);
            }
            lstTablas = lstTablas.OrderBy(x => x.Name).ToList();
            return lstTablas;
        }
        public static List<FillPatternElement> getFillPatternElements(Document doc)
        {
            List<FillPatternElement> fillPatternElements = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).OfType<FillPatternElement>().ToList();
            return fillPatternElements;
        }
        public static FillPatternElement getPatternByName(Document doc, string name)
        {
            List<FillPatternElement> lst = getFillPatternElements(doc);
            FillPatternElement pat = lst.Find(x => x.Name == name);
            return pat;
        }
        /// <summary>  Get the LIst of elements by Category Id </summary>
        public static List<Element> CollectElementsByCategory(Document doc, ElementId categoryId)
        {
            List<Element> lstElement = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> elements = collector.WhereElementIsNotElementType().ToList();

            lstElement = (from elem in elements
                          where elem.Category != null
                          && elem.Category.CategoryType == CategoryType.Model
                          && elem.Category.Id == categoryId
                          select elem).ToList();

            return lstElement;
        }
        public static List<Category> GetModelCategories(Document doc)
        {
            List<Category> lstCat = new List<Category>();
            List<Element> lstElement = new List<Element>();
            Options op = doc.Application.Create.NewGeometryOptions();
            op.ComputeReferences = true;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Element> elements = collector.WhereElementIsNotElementType().ToList();

            lstElement = (from elem in elements
                          where elem.Category != null
                          && elem.get_Geometry(op) != null
                          && elem.CreatedPhaseId.IntegerValue != -1
                          && elem.GetTypeId().IntegerValue != -1
                          && elem.Category.CategoryType == CategoryType.Model
                          && elem.Category.Id != (new ElementId(BuiltInCategory.OST_Cameras))
                          select elem).ToList();

            foreach (Element elem in lstElement)
            {
                if (!lstCat.Exists(x => x.Id == elem.Category.Id))
                {
                    lstCat.Add(elem.Category);
                }
            }
            lstCat = lstCat.OrderBy(x => x.Name).ToList();
            return lstCat;
        }
        /// <summary> Get the ElementId from a Category Name </summary>
        public static ElementId getCategoryIdByName(Document doc, string name)
        {
            List<Category> lst = new List<Category>();
            lst = GetModelCategories(doc);

            return lst.Find(x => x.Name == name).Id;
        }
        /// <summary> Get a list of Unique Names of Family Types by Category </summary>
        public static List<string> getTypesByCategory(Document doc, ElementId categoryId)
        {
            List<string> lst = new List<string>();
            foreach (Element elem in CollectElementsByCategory(doc, categoryId))
            {
                if (!lst.Exists(x => x == elem.Name))
                {
                    lst.Add(elem.Name);
                }
            }
            return lst;
        }
        /// <summary> Create a list of Random Color Red from a list of Family Type Names </summary>
        public static List<System.Drawing.Color> getRandomColorsFromTypes(List<string> lst)
        {
            List<System.Drawing.Color> lstColor = new List<System.Drawing.Color>();
            Random r = new Random();
            foreach (string item in lst)
            {
                System.Drawing.Color color = System.Drawing.Color.FromArgb(r.Next(0, 256), r.Next(0, 256),
                    r.Next(0, 256));
                lstColor.Add(color);
            }
            return lstColor;
        }
        public static Document GetOpenedDocumentFormPath(Autodesk.Revit.ApplicationServices.Application app,
            string path)
        {
            Document doc = null;
            DocumentSet docSet = app.Documents;
            foreach (Document docu in docSet)
            {
                if (docu.PathName == path)
                {
                    doc = docu;
                }
            }
            return doc;
        }
        #endregion

        #region Revit
        /// <summary> Apply a solid pattern with color to a list of elements </summary>
        public static void paintElements(Document doc, List<Element> lst, Color col, ElementId patternId)
        {
            //ElementId SolidId = new ElementId(3); //Solid Pattern
            //Vista
            Autodesk.Revit.DB.View actualView = doc.ActiveView;

            using (Transaction t1 = new Transaction(doc, "Override Color in View"))
            {
                t1.Start();
                //Elemento 1
                OverrideGraphicSettings set1 = new OverrideGraphicSettings();
                set1.SetProjectionFillColor(col);
                set1.SetProjectionFillPatternId(patternId);
                try
                {
                    foreach (Element e in lst)
                    {
                        actualView.SetElementOverrides(e.Id, set1);
                    }
                }
                catch (Exception)
                {
                    t1.RollBack();
                    return;
                }
                t1.Commit();
            }
        }
        public static void paintElement(Document doc, Element elem, System.Drawing.Color color, ElementId patternId)
        {
            //ElementId SolidId = new ElementId(3); //Solid Pattern
            //Vista
            Autodesk.Revit.DB.View actualView = doc.ActiveView;
            
            //Elemento 1
            OverrideGraphicSettings set1 = new OverrideGraphicSettings();
            Color col = new Color(color.R, color.G, color.B);
            set1.SetProjectionFillColor(col);
            set1.SetProjectionFillPatternId(patternId);
            actualView.SetElementOverrides(elem.Id, set1);
        }
        /// <summary> Isolate the elements in specified view </summary>
        public static void isolateElementsView(Document doc, List<ElementId> lst, Autodesk.Revit.DB.View view)
        {
            using (Transaction t = new Transaction(doc, "Isolate Elements in View"))
            {
                t.Start();
                try
                {
                    view.IsolateElementsTemporary(lst);
                }
                catch (Exception)
                {
                    t.RollBack();
                    return;
                }
                t.Commit();
            }
        }
        public static XYZ PickFaceSetWorkPlaneAndPickPoint(UIDocument uidoc, string prompt)
        {
            XYZ point_in_3d = null;

            Document doc = uidoc.Document;

            Reference r = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face, prompt);

            Element e = doc.GetElement(r.ElementId);

            if (null != e)
            {
                PlanarFace face
                  = e.GetGeometryObjectFromReference(r)
                    as PlanarFace;

                if (face != null)
                {
                    Plane plane = Plane.CreateByNormalAndOrigin(face.FaceNormal, face.Origin);

                    Transaction t = new Transaction(doc);

                    t.Start("Temporarily set work plane" + " to pick point in 3D");

                    SketchPlane sp = SketchPlane.Create(doc, plane);

                    uidoc.ActiveView.SketchPlane = sp;
                    uidoc.ActiveView.ShowActiveWorkPlane();

                    try
                    {
                        point_in_3d = uidoc.Selection.PickPoint(prompt);
                    }
                    catch (OperationCanceledException)
                    {
                    }

                    t.RollBack();
                }
            }
            return point_in_3d;
        }
        public static double DistanceBetween2Points(XYZ start, XYZ end)
        {
            return start.DistanceTo(end) * 0.3048;
        }
        #endregion

        #region Excel
        public static List<TreeNode> TablaNodes(List<ViewSchedule> lstTablas)
        {
            List<TreeNode> lst = new List<TreeNode>();
            foreach (ViewSchedule tabla in lstTablas)
            {
                TreeNode node = new TreeNode();
                node.Name = tabla.Name;
                node.Text = tabla.Name;
                lst.Add(node);
            }
            return lst;
        }
        public static List<List<string>> GetData(ViewSchedule viewSchedule)
        {
            List<string> lst = new List<string>();
            //Autodesk.Revit.DB.Element rvtElement = viewSchedule.InternalElement;
            //rvtDb.ViewSchedule viewTable = rvtElement as rvtDb.ViewSchedule;

            TableSectionData table = viewSchedule.GetTableData().GetSectionData(SectionType.Body);
            int nRows = table.NumberOfRows;
            int nColumns = table.NumberOfColumns;

            List<List<string>> dataListRow = new List<List<string>>();
            for (int i = 0; i < nRows; i++)
            {
                List<string> dataListColumn = new List<string>();
                for (int j = 0; j < nColumns; j++)
                {
                    dataListColumn.Add(viewSchedule.GetCellText(SectionType.Body, i, j));
                }
                dataListRow.Add(dataListColumn);
            }
            return dataListRow;
        }
        public static DataTable createDataTable(ViewSchedule tabla)
        {
            List<List<string>> lst = GetData(tabla);
            DataTable dt = new DataTable();
            for (int i = 0; i < lst.Count; i++)
            {
                if (i == 0)
                {
                    foreach (string dato in lst[i])
                    {
                        dt.Columns.Add(dato);
                    }
                }
                else
                {
                    //foreach (string dato in lst[i])
                    //{
                        dt.Rows.Add();
                    //}
                }
            }
            return dt;
        }
        public static XLWorkbook createExcelWb(List<ViewSchedule> lstTablas)
        {
            lstTablas = lstTablas.OrderBy(x => x.Name).ToList();
            XLWorkbook wb = new XLWorkbook();
            int page = 1;
            foreach (ViewSchedule tabla in lstTablas)
            {
                DataTable dt = createDataTable(tabla);
                string sheetName = tabla.Name;
                if (tabla.Name.Count() > 31)
                {
                    sheetName = sheetName.Remove(30);
                }
                wb.Worksheets.Add(dt, sheetName);
                var ws = wb.Worksheet(page);
                page += 1;
                int row = 2;

                List<List<string>> lst = GetData(tabla);
                for (int i = 1; i < lst.Count; i++)
                {
                    for (int j = 0; j < lst[i].Count; j++)
                    {
                        ws.Cell(row, j + 1).Value = lst[i][j];
                    }
                    row += 1;
                }
            }
            return wb;
        }
        #endregion

        #region Encadenados
        public static Element PickElement(UIDocument uidoc)
        {
            //Select an object on screen
            Reference r = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, 
                "Seleccione un elemento");
            //We have picked something
            Element e = uidoc.Document.GetElement(r);
            return e;
        }
        public static List<Element> PickElements(UIDocument uidoc)
        {
            List<Element> lst = new List<Element>();
            List<Reference> lstRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element,
                "Seleccione varios Muros").ToList();
            foreach (Reference re in lstRef)
            {
                Element e = uidoc.Document.GetElement(re);
                lst.Add(e);
            }
            return lst;
        }
        //Moving Wall using Curve
        public void MoveUsingCurveParam(Autodesk.Revit.ApplicationServices.Application application, Wall wall)
        {
            LocationCurve wallLine = wall.Location as LocationCurve;
            XYZ p1 = XYZ.Zero;
            XYZ p2 = new XYZ(10, 20, 0);
            Line newWallLine = Line.CreateBound(p1, p2);

            // Change the wall line to a new line.
            wallLine.Curve = newWallLine;
        }
        //Nombres de las Vigas del Proyecto
        public static List<string> lstArmazon(Document doc)
        {
            List<string> lst = new List<string>();
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming);
            List<Element> elements = collector.WhereElementIsElementType().ToList();

            foreach (Element e in elements)
            {
                FamilySymbol fs = e as FamilySymbol;
                lst.Add(fs.Family.Name + ": " + fs.Name);
            }
            return lst;
        }
        public static List<string> lstLevelNames(Document doc)
        {
            List<string> lst = new List<string>();
            foreach (Level level in getLevels(doc))
            {
                lst.Add(level.Name);
            }
            return lst;
        }
        public static double GetWallHeigth(Document doc, Wall wall)
        {
            double alturaDesconectada = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble() * 0.3048;
            Level level = doc.GetElement(wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId())
                as Level;
            double levelElevation = level.Elevation * 0.3048;
            return alturaDesconectada - levelElevation;
        }
        public static Curve GetCurveFromWall(Wall wall)
        {
            LocationCurve locCurve = wall.Location as LocationCurve;
            Curve cur = locCurve.Curve;
            return cur;
        }
        public static List<FamilySymbol> GetBeamsLoaded(Document doc)
        {
            List<FamilySymbol> lst = new List<FamilySymbol>();
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming);
            List<Element> elements = collector.WhereElementIsElementType().ToList();

            foreach (Element e in elements)
            {
                FamilySymbol fs = e as FamilySymbol;
                lst.Add(fs);
            }
            return lst;
        }
        public static FamilySymbol GetBeamFromName(Document doc, string name)
        {
            FamilySymbol symbol = null;
            foreach (FamilySymbol fs in GetBeamsLoaded(doc))
            {
                string superName = fs.Family.Name + ": " + fs.Name;
                if (superName == name)
                {
                    symbol = fs;
                }
            }
            return symbol;
        }
        #endregion
        
        #region Task Manager
        public static string GetCategoriaFromNode(TreeNode node)
        {
            string[] line = node.Name.Split(';');
            return line[0];
        }
        public static int GetIdFromNode(TreeNode node)
        {
            string[] line = node.Name.Split(';');
            return Convert.ToInt32(line[1]);
        }
        //public static void ExportarTsk(TskProyecto proyecto, string folderPath)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    // Agregar Proyecto
        //    sb.AppendLine(proyecto.ToTskLine());
        //    // Agregar Grupos
        //    List<TskGrupo> lstGrupos = proyecto.GetAllGrupos();
        //    foreach (var item in lstGrupos)
        //    {
        //        sb.AppendLine(item.ToTskLine());
        //    }
        //    // Agregar Tareas
        //    List<TskTarea> lstTareas = proyecto.GetAllTareas();
        //    foreach (var item in lstTareas)
        //    {
        //        sb.AppendLine(item.ToTskLine());
        //    }
        //    // Escribir Archivo Tsk
        //    File.WriteAllText(folderPath, sb.ToString());
        //}
        //public static void ImportarTsk(string path)
        //{
        //    // Crear Proyecto
        //    TskProyecto proyecto = new TskProyecto();
        //    proyecto.Insert();
        //    TskProyecto proyecto0 = TskProyecto.GetLast();
        //    int colorNegro = System.Drawing.Color.Black.ToArgb();

        //    // Listas de objetos de Importación
        //    List<TskGrupoImport> lstGruposImportados = new List<TskGrupoImport>();
        //    List<TskTareaImport> lstTareasImportadas = new List<TskTareaImport>();

        //    // Abrir Archivo Tsk
        //    string[] lines = File.ReadAllLines(path);
        //    for (int i = 0; i < lines.Length; i++)
        //    {
        //        string[] values = lines[i].Split(new string[] { "|||" }, StringSplitOptions.None);
        //        string categoria = values[0];
        //        // Rellenar Listas
        //        switch (categoria)
        //        {
        //            case "Proyecto":
        //                proyecto0.nombre = values[1];
        //                proyecto0.detalles = values[2];
        //                proyecto0.color = Convert.ToInt32(values[3]);
        //                proyecto0.colorF = Convert.ToInt32(values[4]);
        //                proyecto0.Update();
        //                break;
        //            case "Grupo":
        //                TskGrupoImport grupo = new TskGrupoImport(Convert.ToInt32(values[1]), 0, values[2], 
        //                    categoria, Convert.ToInt32(values[3]), proyecto0.id, Convert.ToInt32(values[5]),
        //                    Convert.ToInt32(values[6]), values[4]);
        //                lstGruposImportados.Add(grupo);
        //                break;
        //            case "Tarea":
        //                TskTareaImport tarea = new TskTareaImport(Convert.ToInt32(values[1]), 0, values[2],
        //                    categoria, Convert.ToInt32(values[8]), Convert.ToInt32(values[6]),
        //                    Convert.ToInt32(values[7]), values[3], Convert.ToBoolean(values[4]), 
        //                    proyecto0.id, Convert.ToDecimal(values[5]));
        //                lstTareasImportadas.Add(tarea);
        //                break;
        //        }
        //    }
        //    // Crear Grupos
        //    foreach (var item in lstGruposImportados)
        //    {
        //        TskGrupo grupo = new TskGrupo();
        //        grupo.Insert();
        //        TskGrupo grupo0 = TskGrupo.GetLast();
        //        item.id = grupo0.id;
        //        item.proyectoId = proyecto0.id;
        //        item.UpdateTskGrupo(grupo0);
        //    }
        //    // Crear Tareas
        //    foreach (var item in lstTareasImportadas)
        //    {
        //        TskTarea tarea = new TskTarea();
        //        tarea.Insert();
        //        TskTarea tarea0 = TskTarea.GetLast();
        //        item.id = tarea0.id;
        //        item.proyectoId = proyecto0.id;
        //        item.UpdateTskTarea(tarea0);
        //    }
        //    // Actualizar Parent de Grupos
        //    foreach (var item in lstGruposImportados)
        //    {
        //        TskGrupo grupo0 = TskGrupo.GetById(item.id);
        //        if (item.parentId == 0)
        //        {
        //            grupo0.parentId = 0;
        //        }
        //        else
        //        {
        //            TskGrupoImport parent = lstGruposImportados.FirstOrDefault(x => x.id0 == item.parentId);
        //            TskGrupo parent0 = TskGrupo.GetById(parent.id);
        //            grupo0.parentId = parent0.id;
        //        }
        //        grupo0.Update();
        //    }
        //    // Actualizar Parent de Tareas
        //    foreach (var item in lstTareasImportadas)
        //    {
        //        TskTarea tarea0 = TskTarea.GetById(item.id);
        //        TskGrupoImport parent = lstGruposImportados.FirstOrDefault(x => x.id0 == item.parentId);
        //        TskGrupo parent0 = TskGrupo.GetById(parent.id);
        //        tarea0.parentId = parent0.id;
        //        tarea0.Update();
        //    }
        //}
        public static void CollectExpandedNodes(TreeNode nodes, List<string> lstExpanded)
        {
            foreach (TreeNode node in nodes.Nodes)
            {
                if (node.IsExpanded)
                {
                    lstExpanded.Add(node.Name);
                }
                if (node.Nodes.Count > 0)
                {
                    CollectExpandedNodes(node, lstExpanded);
                }
            }
        }
        #endregion

        #region BenchMark
        public static void CreateSaveAndOpenDocument(UIApplication uiapp, string path)
        {
            Document doc = uiapp.Application.NewProjectDocument(UnitSystem.Metric);
            SaveAsOptions opt = new SaveAsOptions();
            opt.OverwriteExistingFile = true;
            doc.SaveAs(path, opt);
            uiapp.OpenAndActivateDocument(path);
        }
        public static void CreateWalls2(Document doc)
        {
            using (Transaction t = new Transaction(doc, "BenchMark: Create Walls"))
            {
                try
                {
                    t.Start();
                    int total = 50;
                    for (int i = 1; i < total; i++)
                    {
                        XYZ p0 = new XYZ(0, 0, 0);
                        XYZ p1 = new XYZ(2, 0, 0);
                        Line line = Line.CreateBound(p0, p1);
                        Curve cur = line as Curve;
                        Level lvl = getLevels(doc).First();
                        Wall.Create(doc, cur, lvl.Id, false);
                    }

                    t.Commit();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Universo BIM", "Error: " + ex.Message);
                    t.RollBack();
                }
            };
        }
        public static List<Wall> CreateWalls(Document doc, int total)
        {
            List<Wall> walls = new List<Wall>();
            for (int i = 1; i < total; i++)
            {
                for (int j = 1; j < total; j++)
                {
                    XYZ p0 = new XYZ(i / 0.3048, j / 0.3048, 0);
                    XYZ p1 = new XYZ((i + 0.2) / 0.3048, j / 0.3048, 0);
                    XYZ p2 = new XYZ((i + 0.2) / 0.3048, j / 0.3048, 3 / 0.3048);
                    XYZ p3 = new XYZ(i / 0.3048, j / 0.3048, 3 / 0.3048);
                    Line line1 = Line.CreateBound(p0, p1);
                    Line line2 = Line.CreateBound(p1, p2);
                    Line line3 = Line.CreateBound(p2, p3);
                    Line line4 = Line.CreateBound(p3, p0);
                    Curve cur1 = line1 as Curve;
                    Curve cur2 = line2 as Curve;
                    Curve cur3 = line3 as Curve;
                    Curve cur4 = line4 as Curve;
                    List<Curve> lst = new List<Curve>();
                    lst.Add(cur1);
                    lst.Add(cur2);
                    lst.Add(cur3);
                    lst.Add(cur4);
                    Level lvl = getLevels(doc).First();
                    Wall muro = Wall.Create(doc, lst, false);
                    walls.Add(muro);
                }
            }
            return walls;
        }
        public static Group CrearGrupoModelo(Document doc, List<Wall> walls)
        {
            Group grupo = null;
            List<ElementId> ids = new List<ElementId>();
            foreach (Wall muro in walls)
            {
                ids.Add(muro.Id);
            }
            grupo = doc.Create.NewGroup(ids);
            grupo.GroupType.Name = "Muros";
            return grupo;
        }
        public static void TestModelado(Document doc)
        {
            List<Wall> muros = CreateWalls(doc, 25);
            //Group grupo = CrearGrupoModelo(doc, muros);
        }
        public static void CloseActiveDocument(ExternalCommandData commandData)
        {
            commandData.Application.ActiveUIDocument.SaveAndClose();
            //var placeholderFile = "C:/temp/benchmark.rvt";
            //var doc = commandData.Application.ActiveUIDocument.Document;
            //var file = doc.PathName;
            //var docPlaceholder = commandData.Application.OpenAndActivateDocument(placeholderFile);
            //doc.Close(false);
            //var uidoc = commandData.Application.OpenAndActivateDocument(file);
            //docPlaceholder.Document.Close(false);
        }
        #endregion

        #region Sistema
        public static Stopwatch TimerStart()
        {
            Stopwatch reloj = new Stopwatch();
            reloj.Start();
            return reloj;
        }
        public static int TimeGetMiliSeconds(Stopwatch timer)
        {
            int mili;
            timer.Stop();
            long duration = timer.ElapsedMilliseconds;
            mili = Convert.ToInt32(duration);
            return mili;
        }
        public static void MouseWait()
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
        }
        public static void MouseArrow()
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow;
        }
        #endregion

        #region FamilyManager
        public static void FillTreeFamily(FolderBrowserDialog sfd, TreeView tree, List<string> _pathList)
        {
            string[] array = Directory.GetFiles(sfd.SelectedPath, "*.rfa", SearchOption.AllDirectories);
            tree.Nodes.Clear();
            _pathList.Clear();
            foreach (string item in array)
            {
                TreeNode node = new TreeNode();
                node.Name = item;
                node.Text = item;
                tree.Nodes.Add(node);
                _pathList.Add(item);
            }
        }
        public static FamilyInstance CreateModelText(Document doc, string text, XYZ point, Level nivel)
        {
            string familyPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "LG-TextoModelo.rfa");
            doc.LoadFamily(familyPath);
            List<Element> tipos = getFamilies(doc);
            List<FamilySymbol> simbolos = new List<FamilySymbol>();
            foreach (Element elem in tipos)
            {
                FamilySymbol sym = elem as FamilySymbol;
                simbolos.Add(sym);
            }
            FamilySymbol texto = simbolos.FirstOrDefault(x => x.Family.Name == "LG-TextoModelo");
            if (!texto.IsActive)
            {
                texto.Activate();
                doc.Regenerate();
            }
            
            FamilyInstance fam = doc.Create.NewFamilyInstance(point, texto, nivel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            Parameter param = null;
            foreach (Parameter para in fam.Parameters)
            {
                if (para.Definition.Name == "Texto")
                {
                    param = para;
                }
            }
            param.Set(text);
            return fam;
        }
        public static FamilyInstance PlaceInstance(Document doc, FamilySymbol sym, XYZ point, Level nivel)
        {
            if (!sym.IsActive)
            {
                sym.Activate();
                doc.Regenerate();
            }
            return doc.Create.NewFamilyInstance(point, sym, nivel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        }
        #endregion

        #region Sincronizar Ids
        public static void SincronizarIdsEjemplares(Document doc, List<Element> instances)
        {
            using (Transaction t = new Transaction(doc, "Sincronizar Ids Ejemplar"))
            {
                t.Start();
                try
                {
                    foreach (Element elem in instances)
                    {
                        if (elem.GroupId == new ElementId(-1))
                        {
                            foreach (Parameter param in elem.Parameters)
                            {
                                if (param.Definition.Name.ToLower() == "id" && !param.IsReadOnly)
                                {
                                    param.Set(elem.Id.IntegerValue);
                                    break;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Sincronizar Ids", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                t.Commit();
            };

        }
        #endregion

        #region Crear Tablas
        public static void FillSchedule(ViewSchedule tabla, bool itemizar, bool linked, bool mep, bool pipeSystem, bool ductSystem)
        {
            List<SchedulableField> campos = tabla.Definition.GetSchedulableFields().ToList();

            ElementId famAndTypeId = new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM);
            ElementId sizeId = new ElementId(BuiltInParameter.RBS_CALCULATED_SIZE);
            ElementId systemTypeId = new ElementId(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            ElementId ductSystemId = new ElementId(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);

            // Agregar Campos
            ScheduleField fSystemType = null;
            ScheduleField fDuctSystem = null;
            if (pipeSystem)
            {
                fSystemType = tabla.Definition.AddField(ScheduleFieldType.Instance, systemTypeId);
                fSystemType.GridColumnWidth = 0.15;
            }
            if (ductSystem)
            {
                fDuctSystem = tabla.Definition.AddField(ScheduleFieldType.Instance, ductSystemId);
                fDuctSystem.GridColumnWidth = 0.15;
            }
            ScheduleField fFamilyType = tabla.Definition.AddField(ScheduleFieldType.Instance, famAndTypeId);
            fFamilyType.GridColumnWidth = 0.25;
            ScheduleField fSize = null;
            if (mep)
            {
                fSize = tabla.Definition.AddField(ScheduleFieldType.Instance, sizeId);
                fSize.GridColumnWidth = 0.10;
            }
                
            if (CrearTablas._marca)
            {
                ElementId marcaId = new ElementId(BuiltInParameter.DOOR_NUMBER);
                ScheduleField fMarca = tabla.Definition.AddField(ScheduleFieldType.Instance, marcaId);
                fMarca.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
            }
            if (CrearTablas._comentarios)
            {
                ElementId comentariosId = new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                ScheduleField fComentarios = tabla.Definition.AddField(ScheduleFieldType.Instance, comentariosId);
            }
            if (CrearTablas._descripcion)
            {
                ElementId descId = new ElementId(BuiltInParameter.ALL_MODEL_DESCRIPTION);
                ScheduleField fDescripcion = tabla.Definition.AddField(ScheduleFieldType.ElementType, descId);
            }
            if (CrearTablas._modelo)
            {
                ElementId modeloId = new ElementId(BuiltInParameter.ALL_MODEL_MODEL);
                ScheduleField fModelo = tabla.Definition.AddField(ScheduleFieldType.ElementType, modeloId);
            }
            if (CrearTablas._marcaTipo)
            {
                ElementId marcaTipoId = new ElementId(BuiltInParameter.ALL_MODEL_TYPE_MARK);
                ScheduleField fMarcaTipo = tabla.Definition.AddField(ScheduleFieldType.ElementType, marcaTipoId);
                fMarcaTipo.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
            }
            if (!CrearTablas._itemizar)
            {
                ScheduleField fCount = tabla.Definition.AddField(ScheduleFieldType.Count);
                fCount.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                fCount.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            
            // Largo
            ElementId largoId = new ElementId(BuiltInParameter.INSTANCE_LENGTH_PARAM);
            try
            {
                ScheduleField fLength = tabla.Definition.AddField(ScheduleFieldType.Instance, largoId);
                FormatOptions opt = fLength.GetFormatOptions();
                string title = fLength.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fLength.ColumnHeading = title + " [ML]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fLength.SetFormatOptions(opt);
                fLength.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fLength.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception)
            {

            }
            ElementId largoCurvaId = new ElementId(BuiltInParameter.CURVE_ELEM_LENGTH);
            try
            {
                ScheduleField fLength = tabla.Definition.AddField(ScheduleFieldType.Instance, largoCurvaId);
                FormatOptions opt = fLength.GetFormatOptions();
                string title = fLength.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fLength.ColumnHeading = title + " [ML]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fLength.SetFormatOptions(opt);
                fLength.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fLength.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception)
            {

            }
            ElementId largoFootingId = new ElementId(BuiltInParameter.CONTINUOUS_FOOTING_LENGTH);
            try
            {
                ScheduleField fLength = tabla.Definition.AddField(ScheduleFieldType.Instance, largoFootingId);
                FormatOptions opt = fLength.GetFormatOptions();
                string title = fLength.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fLength.ColumnHeading = title + " [ML]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fLength.SetFormatOptions(opt);
                fLength.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fLength.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception)
            {

            }
            // Area
            ElementId areaId = new ElementId(BuiltInParameter.HOST_AREA_COMPUTED);
            try
            {
                ScheduleField fArea = tabla.Definition.AddField(ScheduleFieldType.Instance, areaId);
                FormatOptions opt = fArea.GetFormatOptions();
                string title = fArea.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fArea.ColumnHeading = title + " [M2]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_SQUARE_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fArea.SetFormatOptions(opt);
                fArea.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fArea.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception)
            {

            }
            // Volumen
            ElementId volumeId = new ElementId(BuiltInParameter.HOST_VOLUME_COMPUTED);
            try
            {
                ScheduleField fVolumen = tabla.Definition.AddField(ScheduleFieldType.Instance, volumeId);
                FormatOptions opt = fVolumen.GetFormatOptions();
                string title = fVolumen.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fVolumen.ColumnHeading = title + " [M3]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_CUBIC_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fVolumen.SetFormatOptions(opt);
                fVolumen.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fVolumen.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception ex)
            {

            }

            // Itemizar
            tabla.Definition.IsItemized = itemizar;

            // Agrupar campos
            ScheduleSortGroupField agruparPipeSystem = new ScheduleSortGroupField();
            ScheduleSortGroupField agruparDuctSystem = new ScheduleSortGroupField();
            ScheduleSortGroupField agrupar0 = new ScheduleSortGroupField();
            agrupar0.FieldId = fFamilyType.FieldId;
            ScheduleSortGroupField agrupar1 = new ScheduleSortGroupField(); // For MEP
            if (mep)
            {
                agrupar1.FieldId = fSize.FieldId;
            }
            if (pipeSystem)
            {
                agruparPipeSystem.FieldId = fSystemType.FieldId;
            }
            if (ductSystem)
            {
                agruparDuctSystem.FieldId = fDuctSystem.FieldId;
            }
            if (itemizar)
            {
                agrupar0.ShowBlankLine = true;
            }
            if (pipeSystem)
            {
                agruparPipeSystem.ShowBlankLine = true;
                tabla.Definition.AddSortGroupField(agruparPipeSystem);
            }
            if (ductSystem)
            {
                agruparDuctSystem.ShowBlankLine = true;
                tabla.Definition.AddSortGroupField(agruparDuctSystem);
            }
            tabla.Definition.AddSortGroupField(agrupar0);
            if (mep) // Agrupar para MEP
            {
                tabla.Definition.AddSortGroupField(agrupar1);
            }

            if (linked)
            {
                tabla.Definition.IncludeLinkedFiles = true;
            }

            // Total general
            tabla.Definition.ShowGrandTotal = true;
            tabla.Definition.ShowGrandTotalTitle = true;
        }

        public static void FillScheduleMaterial(ViewSchedule tabla, bool itemizar, bool linked)
        {
            List<SchedulableField> campos = tabla.Definition.GetSchedulableFields().ToList();

            //ElementId famAndTypeId = new ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM);
            ElementId nameId = new ElementId(BuiltInParameter.MATERIAL_NAME);
            ElementId descId = new ElementId(BuiltInParameter.ALL_MODEL_DESCRIPTION);


            // Agregar Campos
            //tabla.Definition.AddField(ScheduleFieldType.Instance, famAndTypeId);
            ScheduleField fMaterialName = tabla.Definition.AddField(ScheduleFieldType.Material, nameId);
            fMaterialName.GridColumnWidth = 0.25;
            ScheduleField fDescription = tabla.Definition.AddField(ScheduleFieldType.Material, descId);
            fDescription.GridColumnWidth = 0.25;
            if (!CrearTablas._itemizar)
            {
                ScheduleField fCount = tabla.Definition.AddField(ScheduleFieldType.Count);
                fCount.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                fCount.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            
            // Area
            ElementId areaId = new ElementId(BuiltInParameter.MATERIAL_AREA);
            try
            {
                ScheduleField fArea = tabla.Definition.AddField(ScheduleFieldType.MaterialQuantity, areaId);
                FormatOptions opt = fArea.GetFormatOptions();
                string title = fArea.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fArea.ColumnHeading = title + " [M2]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_SQUARE_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fArea.SetFormatOptions(opt);
                fArea.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fArea.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception)
            {

            }
            // Volumen
            ElementId volumeId = new ElementId(BuiltInParameter.MATERIAL_VOLUME);
            try
            {
                ScheduleField fVolumen = tabla.Definition.AddField(ScheduleFieldType.MaterialQuantity, volumeId);
                FormatOptions opt = fVolumen.GetFormatOptions();
                string title = fVolumen.ColumnHeading;
                if (!CrearTablas._itemizar)
                {
                    title = "Σ " + title;
                }
                fVolumen.ColumnHeading = title + " [M3]";
                opt.UseDefault = false;
                opt.DisplayUnits = DisplayUnitType.DUT_CUBIC_METERS;
                opt.Accuracy = 0.01;
                opt.UnitSymbol = UnitSymbolType.UST_NONE;
                fVolumen.SetFormatOptions(opt);
                fVolumen.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                fVolumen.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            catch (Exception ex)
            {

            }

            // Itemizar
            tabla.Definition.IsItemized = itemizar;

            // Agrupar campos
            ScheduleSortGroupField agrupar0 = new ScheduleSortGroupField();
            agrupar0.FieldId = fMaterialName.FieldId;
            if (itemizar)
            {
                agrupar0.ShowBlankLine = true;
            }
            tabla.Definition.AddSortGroupField(agrupar0);
            if(linked)
            {
                tabla.Definition.IncludeLinkedFiles = true;
            }

            // Total general
            tabla.Definition.ShowGrandTotal = true;
            tabla.Definition.ShowGrandTotalTitle = true;
        }

        public static void FillTreeSchedules(TreeView tree)
        {
            // TreeNode 0
            TreeNode node0 = new TreeNode();
            node0.Name = "Tablas";
            node0.Text = "Tablas de Cantidades:";
            node0.ImageIndex = 0;
            node0.SelectedImageIndex = 0;
            node0.Checked = true;

            int count = 1;

            // Treenode MAT
            TreeNode nodeMat = new TreeNode();
            nodeMat.Name = "Materiales";
            nodeMat.Text = "Tablas de Materiales:";
            nodeMat.ImageIndex = 0;
            nodeMat.SelectedImageIndex = 0;
            nodeMat.Checked = true;

            int countMats = 1;

            // TreeNode Tabla
            foreach (Category cat in categorias)
            {
                string nombre = "";
                if (CrearTablas._numerar)
                {
                    nombre += count.ToString();
                }
                if (CrearTablas._prefijo != "")
                {
                    nombre += CrearTablas._prefijo;
                }
                nombre += cat.Name;
                if (CrearTablas._sufijo != "")
                {
                    nombre += CrearTablas._sufijo;
                }

                TreeNode node1 = new TreeNode();
                node1.Name = cat.Id.IntegerValue.ToString();
                node1.Text = nombre;
                node1.ImageIndex = 1;
                node1.SelectedImageIndex = 1;
                node1.Checked = true;
                node0.Nodes.Add(node1);
                count++;

                // Add Material Schedules
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
                        nombre = "M";
                        if (CrearTablas._numerar)
                        {
                            nombre += countMats.ToString();// + "m";
                        }
                        if (CrearTablas._prefijo != "")
                        {
                            nombre += CrearTablas._prefijo;
                        }
                        nombre += cat.Name + " (Materiales)";
                        if (CrearTablas._sufijo != "")
                        {
                            nombre += CrearTablas._sufijo;
                        }

                        TreeNode node2 = new TreeNode();
                        node2.Name = cat.Id.IntegerValue.ToString();// + "_M";
                        node2.Text = nombre;
                        node2.ImageIndex = 1;
                        node2.SelectedImageIndex = 1;
                        node2.Checked = true;
                        nodeMat.Nodes.Add(node2);
                        countMats++;
                    }
                }
            }
            node0.Text += "(" + (count-1).ToString() + ")";
            node0.ExpandAll();
            nodeMat.Text += "(" + (countMats - 1).ToString() + ")";
            nodeMat.ExpandAll();
            tree.Nodes.Add(node0);
            tree.Nodes.Add(nodeMat);
        }
        #endregion
    }
}
