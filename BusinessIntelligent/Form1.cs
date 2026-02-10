using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BusinessIntelligent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) 
        {
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Esto crea la ventana clásica de Windows para buscar archivos
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Solo muestra archivos Excel
            openFileDialog.Filter = "Archivos Excel|*.xls;*.xlsx";

            //El titulo de la ventana que se abre al buscar archivos
            openFileDialog.Title = "Seleccionar archivo Excel";

            // Se ejecuta si se presiona en abrir el archivo
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string rutaArchivo = openFileDialog.FileName;
                CargarExcel(rutaArchivo);
            }
        }

        private void CargarExcel(string ruta)
        {
            //Le dice a .NET cómo leer textos del Excel (acentos, ñ, caracteres especiales)
            //Sin esto ExcelDataReader puede fallar
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            //Abre el archivo en modo lectura
            //Using se asegura que se cierre despues
            using (var stream = File.Open(ruta, FileMode.Open, FileAccess.Read))
            {
                //Convierte el archivo Excel en datos que C# pueda entender.
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    //Convertir Excel a DataSet
                    DataSet result = reader.AsDataSet();

                    // Mostrar la primera hoja del Excel
                    dataGridView1.DataSource = result.Tables[0];
                }
            }
        }
    }
}
