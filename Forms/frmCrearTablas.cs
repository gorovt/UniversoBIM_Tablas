using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace UniBimTablas
{
    public partial class frmCrearTablas : Form
    {
        public Autodesk.Revit.DB.Document _doc;

        public frmCrearTablas(Autodesk.Revit.DB.Document doc)
        {
            InitializeComponent();
            this.Text = "UniversoBIM Tablas " + Tools.version;
            _doc = doc;
            // Obtener la lista de categorías del Modelo
            this.txtPrefijo.Text = CrearTablas._prefijo;
            this.txtSufijo.Text = CrearTablas._sufijo;
            this.chkNumerar.Checked = CrearTablas._numerar;
            this.chkItemizar.Checked = CrearTablas._itemizar;
            this.chkMarca.Checked = CrearTablas._marca;
            this.chkComentarios.Checked = CrearTablas._comentarios;
            this.chkDescripcion.Checked = CrearTablas._descripcion;
            this.chkModelo.Checked = CrearTablas._modelo;
            this.chkMarcaTipo.Checked = CrearTablas._marcaTipo;
            Tools.categorias = Tools.GetModelCategories(doc);
            this.treeView1.Nodes.Clear();
            FillTree();
        }

        private void FillTree()
        {
            Tools.FillTreeSchedules(this.treeView1);
        }

        private void btnCrear_Click(object sender, EventArgs e)
        {
            CrearTablas._itemizar = this.chkItemizar.Checked;
            CrearTablas._numerar = this.chkNumerar.Checked;
            CrearTablas._marca = this.chkMarca.Checked;
            CrearTablas._comentarios = this.chkComentarios.Checked;

            if (this.txtPrefijo.Text != "")
            {
                CrearTablas._prefijo = this.txtPrefijo.Text;
            }
            else
            {
                CrearTablas._prefijo = "";
            }
            if (this.txtSufijo.Text != "")
            {
                CrearTablas._sufijo = this.txtSufijo.Text;
            }
            else
            {
                CrearTablas._sufijo = "";
            }
            // Get Selected Categories
            Tools._selectedCategories = new List<Autodesk.Revit.DB.Category>();
            Tools._selectedMatCategories = new List<Autodesk.Revit.DB.Category>();
            // Quantities Schedules
            foreach (TreeNode node in this.treeView1.Nodes)
            {
                foreach (TreeNode node1 in node.Nodes)
                {
                    if (node1.Checked && !node1.Text.Contains("Materiales"))
                    {
                        // Quantity Category
                        int idSel = Convert.ToInt32(node1.Name);
                        Tools._selectedCategories.Add(Tools.categorias.Find(x => x.Id.IntegerValue == idSel));
                    }
                    if (node1.Checked && node1.Text.Contains("Materiales"))
                    {
                        // Material Category
                        int idSel = Convert.ToInt32(node1.Name);
                        Tools._selectedMatCategories.Add(Tools.categorias.Find(x => x.Id.IntegerValue == idSel));
                    }
                }
            }

            this.DialogResult = DialogResult.OK;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.universobim.com/categoria-producto/addin-revit/");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.universobim.com/producto/qex-para-revit-demo/");
        }

        private void chkNumerar_CheckedChanged(object sender, EventArgs e)
        {
            Refresh();
        }

        private void txtPrefijo_TextChanged(object sender, EventArgs e)
        {
            Refresh();
        }

        private void txtSufijo_TextChanged(object sender, EventArgs e)
        {
            Refresh();
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            
            using (PreviewControl preview = new PreviewControl(_doc, _doc.ActiveView.Id))
            {
                using (Form form = new Form())
                {
                    ElementHost elementHost = new ElementHost();
                    elementHost.Location = new System.Drawing.Point(0, 0);

                    elementHost.Dock = DockStyle.Fill;
                    elementHost.TabIndex = 0;
                    elementHost.Parent = form;
                    elementHost.Child = preview;

                    form.ShowDialog();
                }
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            Refresh();
        }

        private void Refresh()
        {
            CrearTablas._prefijo = this.txtPrefijo.Text;
            CrearTablas._sufijo = this.txtSufijo.Text;
            CrearTablas._numerar = this.chkNumerar.Checked;
            CrearTablas._marca = this.chkMarca.Checked;
            CrearTablas._comentarios = this.chkComentarios.Checked;
            CrearTablas._descripcion = this.chkDescripcion.Checked;
            CrearTablas._modelo = this.chkModelo.Checked;
            CrearTablas._marcaTipo = this.chkMarcaTipo.Checked;
            this.treeView1.Nodes.Clear();
            Tools.FillTreeSchedules(this.treeView1);
        }

        private void chkItemizar_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkItemizar.Checked)
            {
                this.grpEjemplar.Enabled = true;
            }
            else
            {
                this.grpEjemplar.Enabled = false;
                CrearTablas._marca = false;
                CrearTablas._comentarios = false;
                this.chkMarca.Checked = false;
                this.chkComentarios.Checked = false;
            }
        }

        private void chkDescripcion_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkDescripcion.Checked)
            {
                CrearTablas._descripcion = true;
            }
            else
            {
                CrearTablas._descripcion = false;
            }
        }

        private void chkModelo_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkModelo.Checked)
            {
                CrearTablas._modelo = true;
            }
            else
            {
                CrearTablas._modelo = false;
            }
        }

        private void chkMarcaTipo_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkMarcaTipo.Checked)
            {
                CrearTablas._marcaTipo = true;
            }
            else
            {
                CrearTablas._marcaTipo = false;
            }
        }

        private void btnTodo_Click(object sender, EventArgs e)
        {
            foreach (TreeNode node0 in this.treeView1.Nodes)
            {
                node0.Checked = true;
                foreach (TreeNode node1 in node0.Nodes)
                {
                    node1.Checked = true;
                }
            }
        }

        private void btnNinguno_Click(object sender, EventArgs e)
        {
            foreach (TreeNode node0 in this.treeView1.Nodes)
            {
                node0.Checked = false;
                foreach (TreeNode node1 in node0.Nodes)
                {
                    node1.Checked = false;
                }
            }
        }
    }
}
