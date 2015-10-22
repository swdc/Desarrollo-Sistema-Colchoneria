namespace prueba1
{
    partial class Frm_Reporte
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource5 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.reportesDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.Rv_Reporte = new Microsoft.Reporting.WinForms.ReportViewer();
            this.Gv_Reporte = new System.Windows.Forms.DataGridView();
            this.Btn_Eliminar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.Cbox_Modulo = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.Btn_Word = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.clienteBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DataSet1 = new prueba1.DataSet1();
            this.clienteTableAdapter = new prueba1.DataSet1TableAdapters.clienteTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.reportesDataSetBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Gv_Reporte)).BeginInit();
            this.panel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.clienteBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // Rv_Reporte
            // 
            reportDataSource5.Name = "DataSet1";
            reportDataSource5.Value = this.reportesDataSetBindingSource;
            this.Rv_Reporte.LocalReport.DataSources.Add(reportDataSource5);
            this.Rv_Reporte.LocalReport.ReportEmbeddedResource = "prueba1.Report1.rdlc";
            this.Rv_Reporte.Location = new System.Drawing.Point(3, 11);
            this.Rv_Reporte.Name = "Rv_Reporte";
            this.Rv_Reporte.Size = new System.Drawing.Size(487, 343);
            this.Rv_Reporte.TabIndex = 1;
            // 
            // Gv_Reporte
            // 
            this.Gv_Reporte.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Gv_Reporte.Location = new System.Drawing.Point(12, 122);
            this.Gv_Reporte.Name = "Gv_Reporte";
            this.Gv_Reporte.ReadOnly = true;
            this.Gv_Reporte.ShowEditingIcon = false;
            this.Gv_Reporte.Size = new System.Drawing.Size(344, 246);
            this.Gv_Reporte.TabIndex = 2;
            this.Gv_Reporte.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // Btn_Eliminar
            // 
            this.Btn_Eliminar.BackColor = System.Drawing.Color.Tomato;
            this.Btn_Eliminar.Location = new System.Drawing.Point(281, 87);
            this.Btn_Eliminar.Name = "Btn_Eliminar";
            this.Btn_Eliminar.Size = new System.Drawing.Size(75, 23);
            this.Btn_Eliminar.TabIndex = 4;
            this.Btn_Eliminar.Text = "Eliminar";
            this.Btn_Eliminar.UseVisualStyleBackColor = false;
            this.Btn_Eliminar.Visible = false;
            this.Btn_Eliminar.Click += new System.EventHandler(this.Btn_Eliminar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 90);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "Reportes";
            // 
            // Cbox_Modulo
            // 
            this.Cbox_Modulo.FormattingEnabled = true;
            this.Cbox_Modulo.Location = new System.Drawing.Point(12, 47);
            this.Cbox_Modulo.Name = "Cbox_Modulo";
            this.Cbox_Modulo.Size = new System.Drawing.Size(121, 21);
            this.Cbox_Modulo.TabIndex = 9;
            this.Cbox_Modulo.SelectedIndexChanged += new System.EventHandler(this.Cbox_Modulo_SelectedIndexChanged);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.SystemColors.Control;
            this.button3.Image = global::prueba1.Properties.Resources.excel_icon;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(496, 258);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(110, 50);
            this.button3.TabIndex = 11;
            this.button3.Text = "Excel";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.Image = global::prueba1.Properties.Resources.pdf_icon;
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(496, 154);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(110, 50);
            this.button2.TabIndex = 10;
            this.button2.Text = "PDF";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Btn_Word
            // 
            this.Btn_Word.BackColor = System.Drawing.SystemColors.Control;
            this.Btn_Word.Image = global::prueba1.Properties.Resources.word_icon;
            this.Btn_Word.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Btn_Word.Location = new System.Drawing.Point(496, 46);
            this.Btn_Word.Name = "Btn_Word";
            this.Btn_Word.Size = new System.Drawing.Size(110, 50);
            this.Btn_Word.TabIndex = 3;
            this.Btn_Word.Text = "Word";
            this.Btn_Word.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Btn_Word.UseVisualStyleBackColor = false;
            this.Btn_Word.Click += new System.EventHandler(this.Btn_Word_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.Rv_Reporte);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.Btn_Word);
            this.panel1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.panel1.Location = new System.Drawing.Point(382, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(616, 367);
            this.panel1.TabIndex = 12;
            this.panel1.Visible = false;
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2});
            this.statusStrip1.Location = new System.Drawing.Point(0, 399);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.statusStrip1.Size = new System.Drawing.Size(370, 22);
            this.statusStrip1.TabIndex = 13;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripStatusLabel1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(119, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(121, 17);
            this.toolStripStatusLabel2.Text = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // clienteBindingSource
            // 
            this.clienteBindingSource.DataMember = "cliente";
            this.clienteBindingSource.DataSource = this.DataSet1;
            // 
            // DataSet1
            // 
            this.DataSet1.DataSetName = "DataSet1";
            this.DataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // clienteTableAdapter
            // 
            this.clienteTableAdapter.ClearBeforeFill = true;
            // 
            // Frm_Reporte
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(370, 421);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.Cbox_Modulo);
            this.Controls.Add(this.Btn_Eliminar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Gv_Reporte);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Frm_Reporte";
            this.Text = "Reporteador";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.reportesDataSetBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Gv_Reporte)).EndInit();
            this.panel1.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.clienteBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DataSet1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer Rv_Reporte;
        private System.Windows.Forms.BindingSource clienteBindingSource;
        private DataSet1 DataSet1;
        private DataSet1TableAdapters.clienteTableAdapter clienteTableAdapter;
        private System.Windows.Forms.BindingSource reportesDataSetBindingSource;
        //private reportesDataSet reportesDataSet;
        private System.Windows.Forms.DataGridView Gv_Reporte;
        private System.Windows.Forms.Button Btn_Word;
        private System.Windows.Forms.Button Btn_Eliminar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox Cbox_Modulo;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.Timer timer1;
    }
}

