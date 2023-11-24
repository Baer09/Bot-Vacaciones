using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Excel;
using SpreadsheetLight;
using System.Collections;
using System.Web;

namespace BOT_2._0
{
    public partial class Form1 : Form
    {
        private string fileLog;
        private string logDir;
        private List<string> LstIdInstances;
        static class Global
        {
            public static string _globalInstance = "";

        }
        public Form1()
        {
            InitializeComponent();
            fileLog = @"c:\bots\log\" + DateTime.Now.ToString("dd MM yyyy") + ".txt";
            logDir = @"c:\bots\log\";
            LstIdInstances = new List<string>();
            string instanciaActual;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            log(fileLog,"L 30 btnCerrar");
            this.Close();
        }

        private void btnGuardarInstancia_Click(object sender, EventArgs e)
        {

            // create folder with idInstancia as a name
            string dir = @"C:\bots\" + txtInstancia.Text;
            Global._globalInstance = txtInstancia.Text;
            //instanciaActual = txtInstancia.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }


            // show number of instances
            txtNoIntancias.Text += System.Environment.NewLine + txtInstancia.Text;
            // txtNoIntancias.Text += System.Environment.NewLine + Global._globalInstance;

            //save instances in an array
            LstIdInstances.Add(txtInstancia.Text);

            txtInstancia.Clear();

            log(fileLog,"L 61 btnInstancia");

        }

        private void txtInstancia_TextChanged(object sender, EventArgs e)
        {


        }

        private void btnIdEmpleado_Click(object sender, EventArgs e)
        {
            //create path
            string instancia = txtInstancia.Text;
            string idEmpleado = lblIdEmpleado.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + idEmpleado + ".txt";
            //create txt file & save data
            //txtInstancia.Text
            savedDataToTxtFile(txtInstancia.Text, lblIdEmpleado.Text, path, txtIdEmpleado.Text, "L 77 btnIdEmpleado");
            txtIdEmpleado.Clear();
        }

        private void btnNombreEmpleado_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string nombreEmpleado = lblNombreEmpleado.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + nombreEmpleado + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblNombreEmpleado.Text, path, txtNombreEmpleado.Text, "L 87 btnNombreEmpleado");
            txtNombreEmpleado.Clear();

        }

        private void btnPlaza_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string plaza = lblPlaza.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + plaza + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblPlaza.Text, path, txtPlaza.Text,"L 98 btnPlaza");
            txtPlaza.Clear();
        }

        private void btnDepartamento_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string departamento = lblDepartamento.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + departamento + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblDepartamento.Text, path, txtDepartamento.Text,"L 108 btnDepartamento");
            txtDepartamento.Clear();
        }

        private void btnGerente_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string gerente = lblGerente.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + gerente + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblGerente.Text, path, txtGerente.Text,"L 118 btnGerente");
            txtGerente.Clear();

        }

        private void btnFechaSolicitud_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string fechaSolicitud = lblFechaSolicitud.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + fechaSolicitud + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblFechaSolicitud.Text, path, txtFechaSolicitud.Text, "L 129 btnFechaSolicitud");
            txtFechaSolicitud.Clear();
        }

        private void btnFechaInicio_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string fechaInicio = lblFechaInicio.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + fechaInicio + ".txt";


            // poner linea mas evento
            // ser mas especifico
            savedDataToTxtFile(txtInstancia.Text, lblFechaInicio.Text, path, txtFechaInicio.Text, "L 141 btnFechaInicio");
            txtFechaInicio.Clear();
        }

        private void btnPeriodoVacaiones_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string periodoVacaciones = lblPeriodoVacaciones.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + periodoVacaciones + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblPeriodoVacaciones.Text, path, txtPeriodoVacaciones.Text, "L 152 btnPeriodoVacaciones ");
            txtPeriodoVacaciones.Clear();


        }

        private void btnDiasSolicitud_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string diasSolicitud = lblDiasSolicitud.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + diasSolicitud + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblDiasSolicitud.Text, path,txtDiasSolicitud.Text , "L 164 btnDiasSolicitud ");
            txtDiasSolicitud .Clear();

        }

        private void btnFechaRetorno_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string fechaRetorno = lblFechaRetorno.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + fechaRetorno + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblFechaRetorno.Text, path, txtFechaRetorno.Text, "L 175 btnFechaRetorno");
            txtFechaRetorno.Clear();

        }

        private void btnEtapaActual_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string etapaAcutal = lblEtapaActual.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + etapaAcutal + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblEtapaActual.Text, path, txtEtapaActual.Text, "L 186 btnEtapaActual");
            txtEtapaActual.Clear();

        }

        private void btnFechaActualizacion_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string fechaActualizacion = lblFechaActualizacion.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + fechaActualizacion + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblFechaActualizacion.Text, path, txtFechaActualizacion.Text, "L 197 btnFechaActualizacion");
            txtFechaActualizacion.Clear();


        }

        private void btnPais_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string pais = lblPais.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + pais + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblPais.Text, path, txtPais.Text, "L 209 btnPais");
            txtPais.Clear();

        }

        public void savedDataToTxtFile(string instancia, string labelName, string path, string txtBox, string message)
        {
            path = path + instancia;
            log(path, txtBox);
            log(fileLog, message + " " + labelName + " " + txtBox);
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            log(fileLog, "L 225 btnGenerar");
            //path
            string path = @"C:\bots\import\";
            string pathTxt = @"C:\bots\";

            //loop all IdInstances in the list 
            foreach (string IdInstance in LstIdInstances)
            {
                // Validate if directory exists
                string folder = path + @"xls";
                string excelPath = folder + @"\" + IdInstance + ".xls";
                if (Directory.Exists(folder))
                {
                    //if the excel exists, delete old version
                    if (File.Exists(excelPath))
                        File.Delete(excelPath);
                }
                else
                    Directory.CreateDirectory(folder);

                // set all values from txt files
                // generar sentencia que todos los string sean del mismo tamano, no genera excel
                // si no cumple las condiciones, dejar en el log que no se completo por la condicion
                // funcion de LOG: (llamar log - reporte), recibe mensaje, id mensaje, linea/unbicacion de mensaje
                // 
                try
                {

                    string[] idEmpleado = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblIdEmpleado.Text + ".txt");
                    string[] nombreEmpleado = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblNombreEmpleado.Text + ".txt");
                    string[] plaza = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblPlaza.Text + ".txt");
                    string[] departamento = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblDepartamento.Text + ".txt");
                    string[] gerente = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblGerente.Text + ".txt");
                    string[] periodoVacaciones = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblPeriodoVacaciones.Text + ".txt");
                    string[] diasSolicitud = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblDiasSolicitud.Text + ".txt");
                    string[] fechaSolicitud = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblFechaSolicitud.Text + ".txt");
                    string[] fechaInicio = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblFechaInicio.Text + ".txt");
                    string[] fechaRetorno = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblFechaRetorno.Text + ".txt");
                    string[] etapaActual = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblEtapaActual.Text + ".txt");
                    string[] fechaActualizacion = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblFechaActualizacion.Text + ".txt");
                    string[] pais = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblPais.Text + ".txt");


                    // adding array into a list
                    List<string[]> listOfArrays = new List<string[]>();
                    listOfArrays.Add(idEmpleado);
                    listOfArrays.Add(nombreEmpleado);
                    listOfArrays.Add(plaza);
                    listOfArrays.Add(departamento);
                    listOfArrays.Add(gerente);
                    listOfArrays.Add(periodoVacaciones);
                    listOfArrays.Add(diasSolicitud);
                    listOfArrays.Add(fechaSolicitud);
                    listOfArrays.Add(fechaInicio);
                    listOfArrays.Add(fechaRetorno);
                    listOfArrays.Add(etapaActual);
                    listOfArrays.Add(fechaActualizacion);
                    listOfArrays.Add(pais);

                    if ((checkLength(listOfArrays) && !checkBlanks(listOfArrays)))
                    {
                        //MessageBox.Show("el tamano de los arreglos son iguales", "length of arrays !",
                        //sMessageBoxButtons.OK, MessageBoxIcon.Error);

                        // Create a new excel from txt files
                        using (SLDocument sl = new SLDocument())
                        {
                            sl.SetCellValue("A1", "IdEmpleado");
                            sl.SetCellValue("B1", "NombreEmpleado");
                            sl.SetCellValue("C1", "Plaza");
                            sl.SetCellValue("D1", "Departamento");
                            sl.SetCellValue("E1", "Gerente");
                            sl.SetCellValue("F1", "PeriodoVacaciones");
                            sl.SetCellValue("G1", "DiasSolicitud");
                            sl.SetCellValue("H1", "FechaSolicitud");
                            sl.SetCellValue("I1", "FechaInicio");
                            sl.SetCellValue("J1", "FechaRetorno");
                            sl.SetCellValue("K1", "EtapaAcutal");
                            sl.SetCellValue("L1", "FechaActualizacion");
                            sl.SetCellValue("M1", "Pais");
                            for (int i = 1; i <= idEmpleado.Length; i++)
                            {
                                // check if an array has en emtpy element or if one of them has a different length

                                sl.SetCellValue(i + 1, 1, idEmpleado[i - 1]);
                                sl.SetCellValue(i + 1, 2, nombreEmpleado[i - 1]);
                                sl.SetCellValue(i + 1, 3, plaza[i - 1]);
                                sl.SetCellValue(i + 1, 4, departamento[i - 1]);
                                sl.SetCellValue(i + 1, 5, gerente[i - 1]);
                                sl.SetCellValue(i + 1, 6, periodoVacaciones[i - 1]);
                                sl.SetCellValue(i + 1, 7, diasSolicitud[i - 1]);
                                sl.SetCellValue(i + 1, 8, fechaSolicitud[i - 1]);
                                sl.SetCellValue(i + 1, 9, fechaInicio[i - 1]);
                                sl.SetCellValue(i + 1, 10, fechaRetorno[i - 1]);
                                sl.SetCellValue(i + 1, 11, etapaActual[i - 1]);
                                sl.SetCellValue(i + 1, 12, fechaActualizacion[i - 1]);
                                sl.SetCellValue(i + 1, 13, pais[i - 1]);
                                
                            }
                            sl.SaveAs(excelPath);
                        }
                    }
                    else if (!checkLength(listOfArrays))
                    {
                        //MessageBox.Show("el tamano de los arreglos no son iguales o txt tiene espacios vacios", "length of arrays !",
                        //MessageBoxButtons.OK, MessageBoxIcon.Error);

                        log(fileLog, "El Tamaño de los arreglos no son iguales");

                    }
                    else if (checkBlanks(listOfArrays))
                    {
                        log(fileLog, "txt tiene espacios vacios");
                    }

                }
                catch (FileNotFoundException ex)
                {
                    log(fileLog, ex.ToString());
                }

                catch (Exception ex)
                {
                    log(fileLog, ex.ToString());
                }
            }

            LstIdInstances = new List<string>();

            this.Close();
        }

        private void txtNoIntancias_TextChanged(object sender, EventArgs e)
        {

        }

        private Boolean checkLength(List<string[]> listOfArrays)
        {
            Boolean sameSize = true;
            int sizeOfArray = listOfArrays[0].Length;
            for (int i = 1; i < listOfArrays.Count(); i++)
            {
                if (sizeOfArray != listOfArrays[i].Length)
                {
                    sameSize = false;
                    break;
                }
            }
            return sameSize;
        }
        private Boolean checkBlanks(List<string[]> listOfArrays)
        {
            Boolean blanks = false;
            for (int i = 0; i < listOfArrays.Count; i++)
            {
                for (int j = 0; j < listOfArrays[i].Length; j++)
                {
                    if (string.IsNullOrWhiteSpace(listOfArrays[i][j]))
                    {
                        blanks = true;
                        break;
                    }
                }
            }
            return blanks;
        }

        private void log(string path, string txtMessage)
        {
            try
            {
                if (!File.Exists(path))
                    File.WriteAllText(path, txtMessage + "\n");
                else
                    File.AppendAllText(path, txtMessage + "\n");
            }
            catch (DirectoryNotFoundException ex)
            {
                log(fileLog, ex.ToString());

                //Console.WriteLine(ex.ToString());
                //MessageBox.Show("The directory does not exist, DirectoryNotFoundException, below the error message: \n\n" + ex.ToString(), "Error Message !",
                //MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch (Exception ex)
            {
                log(fileLog, ex.ToString());
                //Console.WriteLine(ex.ToString());
                //MessageBox.Show("Error when trying to create txt file, below the error message: \n\n" + ex.ToString(), "Error Message !",
                //MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void txtIdEmpleado_TextChanged(object sender, EventArgs e)
        {

        }

       
    }
}