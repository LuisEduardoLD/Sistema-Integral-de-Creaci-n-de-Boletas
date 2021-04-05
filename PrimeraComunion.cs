using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Parroquia_San_Pascual_Bailon
{
    public partial class PrimeraComunion : Form
    {
        public PrimeraComunion()
        {
            InitializeComponent();
            init();
        }

        public void init()
        {            
            string fecha = dateTimePicker1.Value.Day.ToString()
                         + "/" + dateTimePicker1.Value.ToString("MMMM")
                         + "/" + dateTimePicker1.Value.Year.ToString();
            txtFecha.Text = fecha.ToUpper();
            string fecha2 = dateTimePicker2.Value.Day.ToString()
                         + "/" + dateTimePicker2.Value.ToString("MMMM")
                         + "/" + dateTimePicker2.Value.Year.ToString();
            txtFechapc.Text = fecha.ToUpper();
            string fecha3 = dateTimePicker3.Value.Day.ToString()
                         + "/" + dateTimePicker3.Value.ToString("MMMM")
                         + "/" + dateTimePicker3.Value.Year.ToString();
            txtFechaexp.Text = fecha3.ToUpper();

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string fecha = dateTimePicker1.Value.Day.ToString()
                         + "/" + dateTimePicker1.Value.ToString("MMMM")
                         + "/" + dateTimePicker1.Value.Year.ToString();
            txtFecha.Text = fecha.ToUpper();
            
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            string fecha2 = dateTimePicker2.Value.Day.ToString()
                         + "/" + dateTimePicker2.Value.ToString("MMMM")
                         + "/" + dateTimePicker2.Value.Year.ToString();
            txtFechapc.Text = fecha2.ToUpper();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string fecha3 = dateTimePicker3.Value.Day.ToString()
                         + "/" + dateTimePicker3.Value.ToString("MMMM")
                         + "/" + dateTimePicker3.Value.Year.ToString();
            txtFechaexp.Text = fecha3.ToUpper();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> lista = new List<string>();      
            lista.Add(txtNombre.Text);
            lista.Add(txtPadre.Text);
            lista.Add(txtMadre.Text);
            lista.Add(txtPadrino.Text);
            lista.Add(txtLugarBautismo.Text);
            lista.Add(txtLugar.Text);
            lista.Add(txtFecha.Text);
            lista.Add(txtFechapc.Text);
            lista.Add(txtMinistro.Text);
            lista.Add(numLibro.Text);
            lista.Add(numFoja.Text);
            lista.Add(numPartida.Text);
            lista.Add(dateTimePicker3.Value.Day.ToString()); //edia
            lista.Add(dateTimePicker3.Value.ToString("MMMM").ToUpper()); //emes
            lista.Add(dateTimePicker3.Value.Year.ToString()); //eanio
            int opcion = 1;
            Form1.crear(lista,opcion);
            //  printPDFWithAcrobat();
        }

        public void printPDFWithAcrobat()
        {
            string Filepath = (@"C:\Users\lalo\source\repos\Parroquia_San_Pascual_Bailon\segundo.docx");

            using (PrintDialog Dialog = new PrintDialog())
            {
                Dialog.ShowDialog();

                ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                {
                    Verb = "print",
                    CreateNoWindow = true,
                    FileName = Filepath,
                    WindowStyle = ProcessWindowStyle.Minimized
                };

                Process printProcess = new Process();
                printProcess.StartInfo = printProcessInfo;
                printProcess.Start();

                printProcess.WaitForInputIdle();

                Thread.Sleep(3000);

                if (false == printProcess.CloseMainWindow())
                {
                    printProcess.Kill();
                }
            }
        }
    }
}
