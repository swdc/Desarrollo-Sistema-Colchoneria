using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using MySql.Data;
using Microsoft.Reporting.WinForms;
using System.IO;


namespace prueba1
{

    public partial class Frm_Reporte : Form
    {
        public Frm_Reporte()
        {
            InitializeComponent();
        }

        #region Variables
        string sNombreReporteGrid;
        string SNombreReporte;
        int IModulo;
        string SNom_Reporte;
        string SUsuario;
        string SFecha_Hora;
        int Iid_Form;
        string STiempo;
        string ruta;
        string tiporeporte;
        string puntoreporte;
        #endregion

        #region Carga de Form
        public void Form1_Load(object sender, EventArgs e)
        {
            
            Fnc_Cargagrid();
            Cbox_Modulo.Items.Add("reportes");
            Cbox_Modulo.Items.Add("productos");
            Cbox_Modulo.Items.Add("clientes");

            #region Llamada de Variables externas (aun no se usan)
            IModulo = 1;
            SUsuario = "cristhiam";
            SFecha_Hora = DateTime.Now.ToString("G");
            Iid_Form = 1;
            //ruta para guardar los reportes
            ruta = "C:\\Reportes";
            #endregion

            toolStripStatusLabel1.Text = "Usuario: "+SUsuario;
            

        }
        #endregion

        #region Carga de Datos a Reporte
        private void Fnc_Cargareporte()
        {
            // Cbox_Modulo.Text trae el nombre el reporte, tabla y modulo que se reporteara
            // string que trae la tabla para crear el reporte
            string Squery = "Select * from " + Cbox_Modulo.Text + "";
            //LLamada a la dll de conexion
            dll_conexion.Conexion cConectar = new dll_conexion.Conexion();
            cConectar.cLocal();
            MySqlCommand comd = new MySqlCommand(Squery, cConectar.SqlConexion);
            //Se limpia el datasource del reporte
            this.Rv_Reporte.LocalReport.DataSources.Clear();
            Rv_Reporte.Reset();
            //Se llama el reporte que se mostrara en el reportviewer
            Rv_Reporte.LocalReport.ReportEmbeddedResource = "prueba1." + Cbox_Modulo.Text + ".rdlc";
            //se crea el datatable que se llenara
            DataTable Dt_table = new DataTable();
            //Se cargan los valores en dt
            Dt_table.Load(comd.ExecuteReader());
            //Termina la conexion
            cConectar.SqlConexion.Close();
            //Se crea una nueva llamada al reporte por medio del DataSet (se utiliza solo como referencia no conexion)
            ReportDataSource RprtDS_Origen = new ReportDataSource();
            RprtDS_Origen.Name = "DataSet1";
            RprtDS_Origen.Value = Dt_table;
            //Se cargan los datos al reporte
            this.Rv_Reporte.LocalReport.DataSources.Add(RprtDS_Origen);
            //Refresca el reporte
            this.Rv_Reporte.RefreshReport();
        }
        #endregion

        #region Funcion para cargar grid
        private void Fnc_Cargagrid()
        {
            //conexion por dll
            dll_conexion.Conexion cConectar = new dll_conexion.Conexion();
            cConectar.cLocal();
            //query de llamada de datos al grid
            cConectar.sqlData = new MySqlDataAdapter("Select nom_reporte, usuario, fecha_hora from reportes", cConectar.SqlConexion);
            DataTable DT_Table = new DataTable();
            //Carga el grid
            cConectar.sqlData.Fill(DT_Table);
            Gv_Reporte.DataSource = DT_Table;
            //Se renombran los headers de las columnas
            Gv_Reporte.Columns[0].HeaderText = "Reporte";
            Gv_Reporte.Columns[1].HeaderText = "Usuario";
            Gv_Reporte.Columns[2].HeaderText = "Fecha de Creacion";
            //Termina la conexion
            cConectar.SqlConexion.Close();
        }
        #endregion

        #region Seleccion de datos del Grid
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

                //recupera el valor de la columna 0 para ser usado como referencia de nombre
                sNombreReporteGrid = Gv_Reporte[0, Gv_Reporte.CurrentCell.RowIndex].Value.ToString();
                //Abre o ejecuta el pdf llamando con el nombre seleccionado del grid
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = "C:/Reportes/" + sNombreReporteGrid;
                proc.Start();
                proc.Close();
            
        }
        #endregion


        #region Crea Reporte
        private void CreaReporte()
        {
            //variable que contiene el tiempo en formato para guardar en nombre del archivo
            STiempo = DateTime.Now.ToString("yyyyMMddhhmmss");
            //String que lleva la direccion de la carpeta contenedora de los reportes 
            string SSave;
            //Condicion para crear la carpeta en disco c
            //si existe guarda los datos
            //lleva la direccion, nombre, fecha y hora, y el .pdf para crear el documento
            SSave = "C:/Reportes/" + Cbox_Modulo.Text + "_" + STiempo + puntoreporte + "";
            // crea el documento pdf
            byte[] Bytes = Rv_Reporte.LocalReport.Render(format: tiporeporte, deviceInfo: "");
            using (FileStream stream = new FileStream(SSave, FileMode.Create))
                {

                    stream.Write(Bytes, 0, Bytes.Length);
                }
            //variable que lleva el nombre del archivo que se creo pdf.. para la inserccion en BD
            SNom_Reporte = "" + Cbox_Modulo.Text + "_" + STiempo + puntoreporte + "";
            //conexion por dll
            dll_conexion.Conexion cConectar = new dll_conexion.Conexion();
            cConectar.cLocal();
            //query de inserccion
            cConectar.sqlCmd = new MySqlCommand("INSERT INTO reportes (modulo, nom_reporte, usuario, fecha_hora, id_form) VALUES ('" + IModulo + "' , '" + SNom_Reporte + "' , '" + SUsuario + "' , '" + SFecha_Hora + "' , '" + Iid_Form + "')", cConectar.SqlConexion);
            cConectar.sqlCmd.ExecuteNonQuery();
            cConectar.SqlConexion.Close();
            //mensaje de confirmacion
             MessageBox.Show("Reporte Creado");
            //actualiza grid
             this.Size = new Size(386, 421);
            Fnc_Cargagrid();
        }
        #endregion

        #region Llamada a cargar reporte por nombre del box
        private void Cbox_Modulo_SelectedIndexChanged(object sender, EventArgs e)
        {
            Fnc_Cargareporte();
            panel1.Visible = true;
            this.Size = new Size(1007, 421);
        }
        #endregion

        #region Eliminar reporte
        private void Btn_Eliminar_Click(object sender, EventArgs e)
        {
            //llamada a conexion dll
            dll_conexion.Conexion cConectar = new dll_conexion.Conexion();
            cConectar.cLocal();
            //query de eliminacion segun el nombre del archivo o reporte
            cConectar.sqlCmd = new MySqlCommand("DELETE FROM reportes WHERE nom_reporte='" + sNombreReporteGrid + "'", cConectar.SqlConexion);
            cConectar.sqlCmd.ExecuteNonQuery();
            cConectar.SqlConexion.Close();
            Fnc_Cargagrid();
            //metodo de eliminacion de archivos locales
            System.IO.File.Delete("C:/Reportes/" + sNombreReporteGrid);
            //confirmacion de eliminacion
            MessageBox.Show("Registro Eliminado");
        }
        #endregion

        #region Boton Crea Word
        private void Btn_Word_Click(object sender, EventArgs e)
        {
            tiporeporte = "Word";
            puntoreporte = ".doc";
            if (Directory.Exists(ruta))
            {
                CreaReporte();
            }
            else //si no existe
            {
                //crea carpeta para guardar.. y guarda documentos
                Directory.CreateDirectory(ruta);
                CreaReporte();
            }
            //envia el nombre del reporte
            SNombreReporte = Cbox_Modulo.Text;
            
        }
        #endregion

        #region Boton Crea Pdf
        private void button2_Click(object sender, EventArgs e)
        {
            tiporeporte = "PDF";
            puntoreporte = ".pdf";
            if (Directory.Exists(ruta))
            {
                CreaReporte();
            }
            else //si no existe
            {
                //crea carpeta para guardar.. y guarda documentos
                Directory.CreateDirectory(ruta);
                CreaReporte();
            }
            //envia el nombre del reporte
            SNombreReporte = Cbox_Modulo.Text;
        }
        #endregion

        #region Boton Crea Excel
        private void button3_Click(object sender, EventArgs e)
        {
            tiporeporte = "Excel";
            puntoreporte = ".xls";
            if (Directory.Exists(ruta))
            {
                CreaReporte();
            }
            else //si no existe
            {
                //crea carpeta para guardar.. y guarda documentos
                Directory.CreateDirectory(ruta);
                CreaReporte();
            }
            //envia el nombre del reporte
            SNombreReporte = Cbox_Modulo.Text;
        }
        #endregion

        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = "Fecha y Hora: "+DateTime.Now.ToString("G");
        }


    }
}
