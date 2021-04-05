using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Parroquia_San_Pascual_Bailon 
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {

        }

        public static void FindAndReplaceWords(Microsoft.Office.Interop.Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object matchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object matchread_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref matchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        public static void CreateWordDocument(object filename, object SaveAs, List<string> lista, int opcion)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            Microsoft.Office.Interop.Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();
                switch (opcion)
                {
                    case 1:
                        myWordDoc.Bookmarks["nombre"].Range.Text = lista.ElementAt(0);
                        if (lista.ElementAt(1).Contains("-"))
                        {
                            string phrase = lista.ElementAt(1);
                            string[] words = phrase.Split('-');
                            myWordDoc.Bookmarks["nombrespadre"].Range.Text = words[0];
                            myWordDoc.Bookmarks["apaternopadre"].Range.Text = words[1];
                            myWordDoc.Bookmarks["amaternopadre"].Range.Text = words[2];
                        }
                        else
                        {
                            myWordDoc.Bookmarks["nombrespadre"].Range.Text = lista.ElementAt(1);
                            myWordDoc.Bookmarks["apaternopadre"].Range.Text = "";
                            myWordDoc.Bookmarks["amaternopadre"].Range.Text = "";
                        }
                        if (lista.ElementAt(2).Contains("-"))
                        {
                            string phrase = lista.ElementAt(2);
                            string[] words2 = phrase.Split('-');
                            myWordDoc.Bookmarks["nombresmadre"].Range.Text = words2[0];
                            myWordDoc.Bookmarks["apaternomadre"].Range.Text = words2[1];
                            myWordDoc.Bookmarks["amaternomadre"].Range.Text = words2[2];
                        }
                        else
                        {
                            myWordDoc.Bookmarks["nombresmadre"].Range.Text = lista.ElementAt(2);
                            myWordDoc.Bookmarks["apaternomadre"].Range.Text = "";
                            myWordDoc.Bookmarks["amaternomadre"].Range.Text = "";
                        }
                        myWordDoc.Bookmarks["padrino"].Range.Text = lista.ElementAt(3);
                        myWordDoc.Bookmarks["lugarbautismo"].Range.Text = lista.ElementAt(4);
                        myWordDoc.Bookmarks["lugar"].Range.Text = lista.ElementAt(5);
                        myWordDoc.Bookmarks["fecha"].Range.Text = lista.ElementAt(6);
                        myWordDoc.Bookmarks["fechaprimeracom"].Range.Text = lista.ElementAt(7);
                        myWordDoc.Bookmarks["ministro"].Range.Text = lista.ElementAt(8);
                        myWordDoc.Bookmarks["libro"].Range.Text = lista.ElementAt(9);
                        myWordDoc.Bookmarks["foja"].Range.Text = lista.ElementAt(10);
                        myWordDoc.Bookmarks["partida"].Range.Text = lista.ElementAt(11);
                        myWordDoc.Bookmarks["edia"].Range.Text = lista.ElementAt(12);
                        myWordDoc.Bookmarks["emes"].Range.Text = lista.ElementAt(13);
                        myWordDoc.Bookmarks["eanio"].Range.Text = lista.ElementAt(14);
                         
                        break;

                } 

            }
            else
            {
                MessageBox.Show("Archivo no encontrado");
            }


            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                             ref missing, ref missing, ref missing,
                             ref missing, ref missing, ref missing,
                             ref missing, ref missing, ref missing,
                             ref missing, ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            
            MessageBox.Show("Archivo creado");                
        }



        public static void crear(List<string> lista, int opcion)
        {
            switch (opcion)
            {
                case 1:
                    CreateWordDocument(
                        @"C:\Users\lalo\source\repos\Parroquia_San_Pascual_Bailon\BOLETA DE PRIMERA COMUNIÓN.docx",
                        @"C:\Users\lalo\source\repos\Parroquia_San_Pascual_Bailon\segundo.docx", lista, opcion);
                        break;
                case 2:
                    CreateWordDocument(
                        @"C:\Users\lalo\source\repos\Parroquia_San_Pascual_Bailon\platicas prebautismales.docx",
                        @"C:\Users\lalo\source\repos\Parroquia_San_Pascual_Bailon\tercero.docx", lista, opcion);
                        break;
            }
            
            
        }

        public void btnCerrar_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void btnMaximizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btnMaximizar.Visible = false;
            btnRestaurar.Visible = true;
        }

        private void btnRestaurar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            btnRestaurar.Visible = false;
            btnMaximizar.Visible = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void BarraTitulo_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void AbrirFormHijo(object formhijo)
        {
            if (this.PanelContenedor.Controls.Count > 0)
            {
                this.PanelContenedor.Controls.RemoveAt(0);
                Form fh = formhijo as Form;
                fh.TopLevel = false;
                fh.Dock = DockStyle.Fill;
                this.PanelContenedor.Controls.Add(fh);
                this.PanelContenedor.Tag = fh;
                fh.Show();
            }
        }

        private void btnPrimeraComunion_Click(object sender, EventArgs e)
        {
            AbrirFormHijo(new PrimeraComunion());         
        }

        private void btnPlaticasPrebautismales_Click(object sender, EventArgs e)
        {
            AbrirFormHijoPP(new PlaticasPrebautismales());
        }

        private void AbrirFormHijoPP(object formhijoPP)
        {
            if (this.PanelContenedor.Controls.Count > 0)
            {
                this.PanelContenedor.Controls.RemoveAt(0);
                Form fh = formhijoPP as Form;
                fh.TopLevel = false;
                fh.Dock = DockStyle.Fill;
                this.PanelContenedor.Controls.Add(fh);
                this.PanelContenedor.Tag = fh;
                fh.Show();
            }
        }
    }
}
