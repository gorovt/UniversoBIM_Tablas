/*   License
*******************************************************************************
*                                                                             *
*    Copyright (c) 2018-2020 Luciano Gorosito <lucianogorosito@hotmail.com>   *
*                                                                             *
*    This file is part of UniBIM Tablas                                       *
*                                                                             *
*    UniBIM Tablas is free software: you can redistribute it and/or modify    *
*    it under the terms of the GNU General Public License as published by     *
*    the Free Software Foundation, either version 3 of the License, or        *
*    (at your option) any later version.                                      *
*                                                                             *
*    UniBIM Tablas is distributed in the hope that it will be useful,         *
*    but WITHOUT ANY WARRANTY; without even the implied warranty of           *
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the            *
*    GNU General Public License for more details.                             *
*                                                                             *
*    You should have received a copy of the GNU General Public License        *
*    along with this program.  If not, see <https://www.gnu.org/licenses/>.   *
*                                                                             *
*******************************************************************************
*/

using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Novacode;
using ClosedXML.Excel;
using System.Data;

namespace UniBimTablas
{
    public partial class frmExportTablas : System.Windows.Forms.Form
    {
        public frmExportTablas()
        {
            InitializeComponent();
            this.trvTablas.Nodes.Clear();
            fillTree(this.trvTablas, Tools.TablaNodes(Tools.getViewSchedules(ExportarTablas._doc)));
        }

        private void fillTree(TreeView tree, List<TreeNode> lst)
        {
            foreach (TreeNode node in lst)
            {
                if (node.Name.Count() > 31)
                {
                    node.ImageIndex = 1;
                    node.SelectedImageIndex = 1;
                    node.ToolTipText = "Nombre de Tabla muy largo. Max 31 caracteres.";
                }
                tree.Nodes.Add(node);
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            List<ViewSchedule> lstSelected = getSelectedTables(this.trvTablas, Tools.getViewSchedules(ExportarTablas._doc));
            if (lstSelected.Count == 0)
            {
                MessageBox.Show("Debe seleccionar al menos una Tabla", "UniBIM", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else
            {
                string folderPath = string.Empty;
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Documento de Excel|*.xlsx";
                string title = ExportarTablas._doc.Title;
                sfd.FileName = title + ".xlsx";
                DialogResult result = sfd.ShowDialog();
                if (result == DialogResult.OK)
                {
                    folderPath = sfd.FileName;
                    try
                    {
                        XLWorkbook wb = Tools.createExcelWb(lstSelected);
                        wb.SaveAs(folderPath);
                        MessageBox.Show("El documento ha sido creado correctamente", "UniBIM",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("No se puede crear el documento. Detalles: " +
                            ex.Message,
                            "UniBIM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }
        private List<ViewSchedule> getSelectedTables(TreeView tree, List<ViewSchedule> lst)
        {
            List<ViewSchedule> lstTablas = new List<ViewSchedule>();
            foreach (TreeNode node in tree.Nodes)
            {
                if (node.Checked == true)
                {
                    ViewSchedule tabla = lst.Find(x => x.Name == node.Name);
                    lstTablas.Add(tabla);
                }
            }
            return lstTablas;
        }

        private void btnSeleccionarTodo_Click(object sender, EventArgs e)
        {
            foreach (TreeNode node in this.trvTablas.Nodes)
            {
                node.Checked = true;
            }
        }

        private void btnNinguno_Click(object sender, EventArgs e)
        {
            foreach (TreeNode node in this.trvTablas.Nodes)
            {
                node.Checked = false;
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
