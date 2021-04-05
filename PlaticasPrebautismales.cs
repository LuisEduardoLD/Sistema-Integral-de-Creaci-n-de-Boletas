using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Parroquia_San_Pascual_Bailon
{
    public partial class PlaticasPrebautismales : Form
    {
        public PlaticasPrebautismales()
        {
            InitializeComponent();
            init();
        }

        public void init()
        {
            string fecha = dateTimePicker2.Value.Day.ToString()
                         + "/" + dateTimePicker2.Value.ToString("MMMM")
                         + "/" + dateTimePicker2.Value.Year.ToString();
            txtFechaBautismo.Text = fecha.ToUpper();
            string fecha2 = dateTimePicker3.Value.Day.ToString()
                         + "/" + dateTimePicker3.Value.ToString("MMMM")
                         + "/" + dateTimePicker3.Value.Year.ToString();
            txtPresente.Text = fecha.ToUpper();
           

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            string fecha2 = dateTimePicker2.Value.Day.ToString()
                         + "/" + dateTimePicker2.Value.ToString("MMMM")
                         + "/" + dateTimePicker2.Value.Year.ToString();
            txtFechaBautismo.Text = fecha2.ToUpper();
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            string fecha3 = dateTimePicker3.Value.Day.ToString()
                         + "/" + dateTimePicker3.Value.ToString("MMMM")
                         + "/" + dateTimePicker3.Value.Year.ToString();
            txtPresente.Text = fecha3.ToUpper();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> lista = new List<string>();
            lista.Add(txtSr.Text);
            lista.Add(txtNomBau.Text);
            lista.Add(txtCiudad.Text);
            lista.Add(dateTimePicker2.Value.Day.ToString()); 
            lista.Add(dateTimePicker2.Value.ToString("MMMM").ToUpper()); 
            lista.Add(dateTimePicker2.Value.Year.ToString()); 
            lista.Add(txtFechaBautismo.Text);
            lista.Add(txtNombrePadre.Text);
            lista.Add(txtNombreMadre.Text);
            lista.Add(txtNombrePadrino.Text);
            lista.Add(txtNombreMadrina.Text);
            lista.Add(numLibro.Text);
            lista.Add(numFoja.Text);
            lista.Add(numPartida.Text);
            lista.Add(txtPresente.Text);
            lista.Add(dateTimePicker3.Value.Day.ToString()); 
            lista.Add(dateTimePicker3.Value.ToString("MMMM").ToUpper()); 
            lista.Add(dateTimePicker3.Value.Year.ToString());
            int opcion = 2;
            Form1.crear(lista, opcion);
        }
    }
}
