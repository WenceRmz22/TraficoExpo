using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using EPPlus;
using System.Xml;
using OfficeOpenXml;

namespace ReportesTraficoExpo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string name = "";

        private void groupBox1_Enter(object sender, EventArgs e)
        {
            
        }
        public void readExcel(string urlExcel,string nameExcel)
        {
            FileInfo doc = new FileInfo(urlExcel);
            using (ExcelPackage excel = new ExcelPackage(doc))
            {
                var teacherWorksheet = excel.Workbook.Worksheets.Single(ws => ws.Name == nameExcel);
                var cells = teacherWorksheet.Cells;
                int rowCount = cells["A:A"].Count();
                for (int i = 2; i <= rowCount; i++)
                {
                   
                    string nameCove = @""+txtDestino.Text+"\\COVE" + cells["E" + i].Value.ToString() + ".xml";
                    string fechaConGuion = cells["A" + i].Value.ToString().Replace(".", "-");
                    XmlTextWriter writer = new XmlTextWriter(nameCove, System.Text.Encoding.UTF8);
                    writer.WriteStartDocument(true);
                    writer.Formatting = Formatting.Indented;
                    writer.Indentation = 2;
                    writer.WriteStartElement("solicitarRecibirCoveServicio");
                    writer.WriteStartElement("comprobantes");
                    writer.WriteStartElement("C601PATEN");
                    writer.WriteString("3649");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C601ADUSEC");
                    writer.WriteString("270");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C601TIPOPE");
                    writer.WriteString("TOCE.EXP");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C602PATEN");
                    writer.WriteString("3649");
                    writer.WriteEndElement();
                    writer.WriteStartElement("D601FECEXP");
                    writer.WriteString("");
                    writer.WriteEndElement();
                    writer.WriteStartElement("M601OBSERV");
                    writer.WriteString("");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C603RFC");
                    writer.WriteString("CCO920213F84");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C601TIPFIG");
                    writer.WriteString("1");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C601EMAIL");
                    writer.WriteString("traficoexpo@tramitaciones.com");
                    writer.WriteEndElement();
                    writer.WriteStartElement("C601FACDORI");
                    writer.WriteString(cells["E" + i].Value.ToString());
                    writer.WriteEndElement();
                    createNodeFactura(cells["V"+i].Value.ToString().Replace(",","."), "1.0000000", cells["V" + i].Value.ToString().Replace(",", "."), "MEX", writer);
                    createNodeEmisor(writer);
                    createNodeDestinatario(writer);
                    createNodeMercancia(cells["P" + i].Value.ToString(), "BX", "USD", cells["Q" + i].Value.ToString(), cells["R" + i].Value.ToString().Replace(",", "."), cells["V" + i].Value.ToString().Replace(",", "."), cells["V" + i].Value.ToString().Replace(",", "."), writer);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                    writer.Close();
                
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //readExcel();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
          
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }
        private void createNodeMercanciaDescripcion(XmlTextWriter writer)
        {
            writer.WriteStartElement("descripcionesEspecificas");
            writer.WriteStartElement("C606MARCA");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C606MODELO");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C606SUBMOD");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C606NUMSER");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        private void createNodeMercancia(string M605DSCMCIA, string C605UNIUMC,string C605TIPMON,string F605CANUMC,string F605VALUNI,string F605VALTOT,string F605VALDOL, XmlTextWriter writer)
        {
            writer.WriteStartElement("mercancias");
            writer.WriteStartElement("M605DSCMCIA");
            writer.WriteString(M605DSCMCIA);
            writer.WriteEndElement();
            writer.WriteStartElement("C605UNIUMC");
            writer.WriteString(C605UNIUMC);
            writer.WriteEndElement();
            writer.WriteStartElement("C605TIPMON");
            writer.WriteString(C605TIPMON);
            writer.WriteEndElement();
            writer.WriteStartElement("F605CANUMC");
            writer.WriteString(F605CANUMC);
            writer.WriteEndElement();
            writer.WriteStartElement("F605VALUNI");
            writer.WriteString(F605VALUNI);
            writer.WriteEndElement();
            writer.WriteStartElement("F605VALTOT");
            writer.WriteString(F605VALTOT);
            writer.WriteEndElement();
            writer.WriteStartElement("F605VALDOL");
            writer.WriteString(F605VALDOL);
            writer.WriteEndElement();
            createNodeMercanciaDescripcion(writer);
            writer.WriteEndElement();
        }
        private void createNodeDestinatario(XmlTextWriter writer)
        {
            writer.WriteStartElement("destinatario");
            writer.WriteStartElement("C608CVEAM3");
            writer.WriteString("CORONA");
            writer.WriteEndElement();
            writer.WriteStartElement("I608TIPIDE");
            writer.WriteString("0");
            writer.WriteEndElement();
            writer.WriteStartElement("C608IDENTIF");
            writer.WriteString("20-5300132");
            writer.WriteEndElement();
            writer.WriteStartElement("C608APPAT");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C608APMAT");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C608NOM");
            writer.WriteString("CROWN IMPORTS LLC");
            writer.WriteEndElement();
            createNodeDomicilioDest(writer);
            writer.WriteEndElement();
        }
        private void createNodeDomicilioDest(XmlTextWriter writer)
        {
            writer.WriteStartElement("domicilio");
            writer.WriteStartElement("C608CALLE");
            writer.WriteString("SOUTH DEARBORN");
            writer.WriteEndElement();
            writer.WriteStartElement("C608NUMEXT");
            writer.WriteString("131");
            writer.WriteEndElement();
            writer.WriteStartElement("C608NUMINT");
            writer.WriteString("SUITE 1200");
            writer.WriteEndElement();
            writer.WriteStartElement("C608COLONIA");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C608LOCALI");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C608MCPIO");
            writer.WriteString("CHICAGO");
            writer.WriteEndElement();
            writer.WriteStartElement("C608ESTADO");
            writer.WriteString("IL");
            writer.WriteEndElement();
            writer.WriteStartElement("C608PAIS");
            writer.WriteString("USA");
            writer.WriteEndElement();
            writer.WriteStartElement("C608CODPOS");
            writer.WriteString("60603");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        private void createNodeEmisor(XmlTextWriter writer)
        {
            writer.WriteStartElement("emisor");
            writer.WriteStartElement("C607CVEAM3");
            writer.WriteString("C.NAVA");
            writer.WriteEndElement();
            writer.WriteStartElement("I607TIPIDE");
            writer.WriteString("1");
            writer.WriteEndElement();
            writer.WriteStartElement("C607IDENTIF");
            writer.WriteString("CCO920213F84");
            writer.WriteEndElement();
            writer.WriteStartElement("C607APPAT");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C607APMAT");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C607NOM");
            writer.WriteString("COMPANIA CERVECERA DE COAHUILA S. DE R.L. DE C.V.");
            writer.WriteEndElement();
           createNodeDomicilio(writer);
            writer.WriteEndElement();
        }
        private void createNodeDomicilio(XmlTextWriter writer)
        {
            writer.WriteStartElement("domicilio");
            writer.WriteStartElement("C607CALLE");
            writer.WriteString("CARRETERA 57 KM 233.2");
            writer.WriteEndElement();
            writer.WriteStartElement("C607NUMEXT");
            writer.WriteString("85");
            writer.WriteEndElement();
            writer.WriteStartElement("C607NUMINT");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C607COLONIA");
            writer.WriteString("NAVA");
            writer.WriteEndElement();
            writer.WriteStartElement("C607LOCALI");
            writer.WriteString("");
            writer.WriteEndElement();
            writer.WriteStartElement("C607MCPIO");
            writer.WriteString("NAVA");
            writer.WriteEndElement();
            writer.WriteStartElement("C607ESTADO");
            writer.WriteString("CO");
            writer.WriteEndElement();
            writer.WriteStartElement("C607PAIS");
            writer.WriteString("MEX");
            writer.WriteEndElement();
            writer.WriteStartElement("C607CODPOS");
            writer.WriteString("26170");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        private void createNodeFactura(string F604VALMEX, string F604FACMEX, string F604VALDOL,string C604PAIS, XmlTextWriter writer)
        {
            writer.WriteStartElement("factura");
            writer.WriteStartElement("I604CERTORI");
            writer.WriteString("0");
            writer.WriteEndElement();
            writer.WriteStartElement("I604SUBDIV");
            writer.WriteString("0");
            writer.WriteEndElement();
            writer.WriteStartElement("C604CVEINC");
            writer.WriteString("EXW");
            writer.WriteEndElement();
            writer.WriteStartElement("C604VINCU");
            writer.WriteString("S");
            writer.WriteEndElement();
            writer.WriteStartElement("C604MONFAC");
            writer.WriteString("USD");
            writer.WriteEndElement();
            writer.WriteStartElement("F604VALMEX");
            writer.WriteString(F604VALMEX);
            writer.WriteEndElement();
            writer.WriteStartElement("F604FACMEX");
            writer.WriteString(F604FACMEX);
            writer.WriteEndElement();
            writer.WriteStartElement("F604VALDOL");
            writer.WriteString(F604VALDOL);
            writer.WriteEndElement();
            writer.WriteStartElement("C604PAIS");
            writer.WriteString(C604PAIS);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private void txtDestino_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (txtArchivo.Text == "")
                {
                    throw new Exception("Debe ingresar un archivo");
                }
                else
                {
                    if (txtDestino.Text == "")
                    {
                        throw new Exception("Debe ingresar una direccion de destino");
                    }
                    else
                    {
                        string url = txtArchivo.Text;
                        string[] partUrl = url.Split('\\');
                        string[] onlyName = partUrl[partUrl.Length - 1].Split('.');
                        name = onlyName[0];
                        readExcel(url, name);

                        MessageBox.Show("Proceso terminado ! ");
                    }
                }
            }
            catch (Exception ep)
            {
                MessageBox.Show(ep.Message);
            }
            
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@""+txtDestino.Text);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Worksheets 2003(*.xls)|*.xlsx|Excel Worksheets 2007(*.xlsx)|*.xls";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtArchivo.Text = openFileDialog1.FileName;

            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    txtDestino.Text = fbd.SelectedPath.ToString();
                }
            }
        }
    }
}
