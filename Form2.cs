using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPrueba1_11_10_2019
{
    public partial class Form2 : Form
    {
        prueba33 F1 = new prueba33();
        Valvula_de_Venteo_de_Baja F2 = new Valvula_de_Venteo_de_Baja();
        Valvula_Dump F3 = new Valvula_Dump();
        public Form2()
        {
            InitializeComponent();
        }

        private void PRUEBA1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            F1.ShowDialog();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void VálvulaDeVenteoDeBajoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            F3.ShowDialog();
        }
        private void VálvulaDumppToolStripMenuItem_Click(object sender, EventArgs e)
        {
            F2.ShowDialog();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dialogo = MessageBox.Show("¿Desea cerrar el programa?",
                  "Cerrar el programa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogo == DialogResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }
    }
}
