using ComponentFactory.Krypton.Toolkit;
using Squirrel;
using System;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SR
{
    public partial class LibroVentas : KryptonForm
    {
        #region Constructor por defecto
        public LibroVentas()
        {
            InitializeComponent();
        }
        #endregion
        #region Método de verificación de actualizaciones
        private async void CheckForUpdate()
        {
            try
            {
                using (var mgr = await UpdateManager.GitHubUpdateManager("https://github.com/Ficss/SistemaRecaudacion"))
                {
                    try
                    {
                        var updateInfo = await mgr.CheckForUpdate();
                        if (updateInfo.ReleasesToApply.Any())
                        {
                            var versionCount = updateInfo.ReleasesToApply.Count;
                            MessageBox.Show($"{versionCount} actualización encontrada");
                            var versionWord = versionCount > 1 ? "versiones" : "version";
                            var message = new StringBuilder().AppendLine($"La aplicación está {versionCount} {versionWord} detrás.").
                                                              AppendLine("Si elige actualizar, los cambios no tomarán efectos hasta que la aplicación no sea reiniciada.").
                                                              AppendLine("¿Desea descargar e instalar la actualización?").
                                                              AppendLine("Ante cualquier duda llamar al anexo 219 o 220").
                                                              ToString();
                            var result = MessageBox.Show(message, "¿Actualizar Aplicación?", MessageBoxButtons.YesNo);
                            if (result != DialogResult.Yes)
                            {
                                notificacion("Actualización rechazada por el usuario");
                                return;
                            }
                            notificacion("Descargando actualización");
                            var updateResult = await mgr.UpdateApp();

                            notificacion($"Descarga completa. Versión {updateResult.Version} tomará efecto cuando la aplicación sea reiniciada.");
                        }
                        else
                        {
                            notificacionInicio.BalloonTipIcon = ToolTipIcon.Info;
                            notificacionInicio.BalloonTipTitle = "Carta de morosidad";
                            notificacionInicio.BalloonTipText = "No hay actualizaciones pendientes";
                            notificacionInicio.ShowBalloonTip(5000);
                        }
                    }
                    catch (Exception ex)
                    {
                        notificacionInicio.BalloonTipIcon = ToolTipIcon.Warning;
                        notificacionInicio.BalloonTipTitle = "Carta de morosidad";
                        notificacionInicio.BalloonTipText = $"¡Hubo un problema durante el proceso de actualización! {ex.Message}";
                        notificacionInicio.ShowBalloonTip(5000);
                    }
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message + Environment.NewLine;
                if (ex.InnerException != null)
                    message += ex.InnerException.Message;
                MessageBox.Show(message);
            }
        }
        #endregion
        #region Método load
        private void LibroVentas_Load(object sender, EventArgs e)
        {
            try
            {
                CheckForUpdate();
                kryptonPage2.Enabled = false;
                kryptonPage5.Enabled = false;
                btnSiguiente.Enabled = false;
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Recuperar Datos de Servicios Facturados
        private void btnRecuperarServicios_Click(object sender, EventArgs e)
        {
            try
            {
                Stopwatch watch = Stopwatch.StartNew();
                DateTime fecha_inicial = dtpInicio.Value.Date;
                DateTime fecha_final = dtpFinal.Value.Date;
                string fecha_i = fecha_inicial.ToString("dd/MM/yyyy");
                string fecha_f = fecha_final.ToString("dd/MM/yyyy");

                string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
                using (OdbcConnection connection = new OdbcConnection(con))
                {
                    using (OdbcCommand cmd = new OdbcCommand("SELECT m.num_e as FOLIO, m.tipodte as 'TIPO DTE', m.fec_e as FECHA, cast(c.rut as varchar) + '-' + cast(c.dig as varchar) as RUT, RTRIM(c.razon) as RAZÓN, d.item as 'CONCEPTO FACTURA', d.p_real as VALOR, a.codigo as 'CÓDIGO SERVICIO', a.nombre as 'NOMBRE SERVICIO', a.cta_softland as 'CUENTA SOFTLAND' " +
                                                           "FROM movi_enc m " +
                                                           "INNER JOIN clientes c ON c.rut = m.rut_e " +
                                                           "INNER JOIN movi_det d ON d.num_e = m.num_e AND d.tipodte = m.tipodte " +
                                                           "INNER JOIN articulo a ON a.codigo = d.cod_d " +
                                                           "WHERE m.fec_e >= ? " +
                                                           "AND m.fec_e <= ? " +
                                                           "AND m.est_e <> 8 " +
                                                           "order by m.tipodte, m.num_e, d.item ", connection))
                    {
                        cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                        cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();

                        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvServicios.DataSource = dt;
                        }
                        watch.Stop();
                        var tiempo = watch.Elapsed;
                        kryptonLabel4.Text = "Segundos transcurridos: "+tiempo;
                        btnSiguiente.Enabled = true;
                        Afecta(fecha_inicial, fecha_final);
                        Exenta(fecha_inicial, fecha_final);
                        NC(fecha_inicial, fecha_final);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Botón Siguiente para habilitar datos de libro de ventas
        private void btnSiguiente_Click(object sender, EventArgs e)
        {
            Stopwatch watch = Stopwatch.StartNew();
            kryptonPage2.Enabled = true;
            kryptonNavigator1.SelectedPage = kryptonPage2;
            try
            {
                for (int i = 0; i < dgvServicios.Rows.Count; i++)
                {
                    if (dgvServicios.Rows[i].Cells[9].Value.ToString() == string.Empty )
                    {
                        MessageBox.Show("FALTA UNA CUENTA DE SOFTLAND, VERIFIQUE");
                    }
                    break;
                }


                DateTime fecha_inicial = dtpInicio.Value.Date;
                DateTime fecha_final = dtpFinal.Value.Date;
                string fecha_i = fecha_inicial.ToString("dd/MM/yyyy");
                string fecha_f = fecha_final.ToString("dd/MM/yyyy");

                string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
                using (OdbcConnection connection = new OdbcConnection(con))
                {
                    using (OdbcCommand cmd = new OdbcCommand("select FOLIO=m.num_e,'TIPO DTE'=m.tipodte,FECHA=m.fec_e,RUT=convert(varchar,c.rut)+'-'+convert(varchar,c.dig),CLIENTE=TRIM(c.razon), CASE WHEN m.tipodte = 33 then 0 else CONVERT(integer, m.mon_e) end as EXCENTO, CASE WHEN m.tipodte = 34 then 0 WHEN m.tipodte = 61 then 0 else m.net_e end as AFECTO, m.iva_e AS IVA, CONVERT(integer, m.mon_e) AS TOTAL, n.num_fact as FACTURA, a.nombre as 'NOMBRE SERVICIO' " +
                        "from movi_enc as m join " +
                        "clientes as c on c.rut = m.rut_e " +
                        "left join ncre_fac as n on n.num_ncre = m.num_e and n.tiponc = m.tipodte and n.tiponc= 61 " +
                        "INNER JOIN movi_det d ON d.num_e = m.num_e AND d.tipodte = m.tipodte AND d.item = 1 " + 
                        "INNER JOIN articulo a ON a.codigo = d.cod_d " +
                        "where m.fec_e >= ? " +
                        "and m.fec_e <= ? " +
                        "and m.est_e <> 8 " +
                        "order by m.tipodte asc, m.num_e asc", connection))
                    {
                        cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                        cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();

                        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvLibro.DataSource = dt;
                        }
                        btnSiguiente.Enabled = true;
                        watch.Stop();
                        var tiempo = watch.Elapsed;
                        kryptonLabel5.Text = "Segundos transcurridos: " + tiempo;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Carga cuentas de ingreso afecta
        private void Afecta(DateTime fecha_inicial, DateTime fecha_final) {
            string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
            using (OdbcConnection connection = new OdbcConnection(con))
            {
                using (OdbcCommand cmd = new OdbcCommand("select distinct(a.cta_softland) as 'CTA INGRESO SOFTLAND' " +
                    "from movi_enc as m " +
                    "join clientes as c on c.rut = m.rut_e " +
                    "join movi_det as d on d.num_e = m.num_e and d.tipodte = m.tipodte " +
                    "join articulo as a on a.codigo = d.cod_d " +
                    "where m.fec_e >= ? " +
                    "and m.fec_e <= ? " +
                    "and m.est_e <> 8 " +
                    "and m.tipodte = 33", connection))
                {
                    cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                    cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                    connection.Open();
                    cmd.ExecuteNonQuery();
                    connection.Close();

                    using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dgvAfectas.DataSource = dt;
                    }
                }
            }
        }
        #endregion
        #region Carga cuentas de ingreso exenta
        private void Exenta(DateTime fecha_inicial, DateTime fecha_final) {
            string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
            using (OdbcConnection connection = new OdbcConnection(con))
            {
                using (OdbcCommand cmd = new OdbcCommand("select distinct(a.cta_softland) as 'CTA INGRESO SOFTLAND' " +
                    "from movi_enc as m " +
                    "join clientes as c on c.rut = m.rut_e " +
                    "join movi_det as d on d.num_e = m.num_e and d.tipodte = m.tipodte " +
                    "join articulo as a on a.codigo = d.cod_d " +
                    "where m.fec_e >= ? " +
                    "and m.fec_e <= ? " +
                    "and m.est_e <> 8 " +
                    "and m.tipodte = 34", connection))
                {
                    cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                    cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                    connection.Open();
                    cmd.ExecuteNonQuery();
                    connection.Close();

                    using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dgvExentas.DataSource = dt;
                    }
                }
            }
        }
        #endregion
        #region Carga cuentas de ingreso de notas de créditos
        private void NC(DateTime fecha_inicial, DateTime fecha_final) {
            string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
            using (OdbcConnection connection = new OdbcConnection(con))
            {
                using (OdbcCommand cmd = new OdbcCommand("select distinct(a.cta_softland) as 'CTA INGRESO SOFTLAND' " +
                    "from movi_enc as m " +
                    "join clientes as c on c.rut = m.rut_e " +
                    "join movi_det as d on d.num_e = m.num_e and d.tipodte = m.tipodte " +
                    "join articulo as a on a.codigo = d.cod_d " +
                    "where m.fec_e >= ? " +
                    "and m.fec_e <= ? " +
                    "and m.est_e <> 8 " +
                    "and m.tipodte = 61", connection))
                {
                    cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                    cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                    connection.Open();
                    cmd.ExecuteNonQuery();
                    connection.Close();

                    using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        dgvNC.DataSource = dt;
                    }
                }
            }
        }
        #endregion
        #region Botón siguiente para habilitar datos de libro base
        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            DateTime fecha_inicial = dtpInicio.Value.Date;
            DateTime fecha_final = dtpFinal.Value.Date;
            dgvClientes.Rows.Clear();
            dgvIngresos.Rows.Clear();
            TablaCLientes();
            TablaIVA(fecha_inicial, fecha_final);
            TablaIngreso(fecha_inicial, fecha_final);
            kryptonPage5.Enabled = true;
            kryptonNavigator1.SelectedPage = kryptonPage5;
        }
        #endregion
        #region Llenar tabla de clientes
        private void TablaCLientes()
        {
            try
            {
                dgvClientes.Rows.Clear();
                string mes = dtpFinal.Value.ToString("MMMM").ToUpper();
                string cuenta = "1-1-03-01-01";
                string tipodcto = null;
                string tipodctoref = null;
                string glosanc = null;
                string numdctoref = null;
                for (int i = 0; i < dgvLibro.Rows.Count; i++)
                {
                    dgvClientes.Rows.Add();
                    //CUENTAS CLIENTES
                    dgvClientes.Rows[i].Cells[0].Value = cuenta;
                    //ORDENA MONTOS A LA CUENTA DEBE O HABER DEPENDIENDO DEL TIPO DE DOCUMENTO
                    if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("33") || dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("34"))
                    {
                        //MONTO AL DEBE
                        dgvClientes.Rows[i].Cells[1].Value = dgvLibro.Rows[i].Cells[8].Value.ToString();
                    }
                    else
                    {
                        //MONTO AL HABER
                        dgvClientes.Rows[i].Cells[2].Value = dgvLibro.Rows[i].Cells[8].Value.ToString();
                    }
                    //GLOSA
                    if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("61"))
                    {
                        glosanc = "REBAJA FX - " + dgvLibro.Rows[i].Cells[9].Value.ToString() + " POR " + dgvLibro.Rows[i].Cells[10].Value.ToString();
                    }
                    else
                    {
                        glosanc = "FACTURA " + mes + " " + dgvLibro.Rows[i].Cells[10].Value.ToString();
                    }

                    if (glosanc.Length < 60)
                    {
                        dgvClientes.Rows[i].Cells[3].Value = glosanc;
                    }
                    else
                    {
                        dgvClientes.Rows[i].Cells[3].Value = glosanc.Substring(0, 59);
                    }
                    #region vacio 1
                    //EQUIVALENCIA (VACIO)
                    dgvClientes.Rows[i].Cells[4].Value = string.Empty;
                    //MONTO AL DEBE MONEDA ADICIONAL (VACIO)
                    dgvClientes.Rows[i].Cells[5].Value = string.Empty;
                    //MONTO AL HABER MONEDA ADICIONAL (VACIO)
                    dgvClientes.Rows[i].Cells[6].Value = string.Empty;
                    //CODIGO CONDICION VENTA (VACIO)
                    dgvClientes.Rows[i].Cells[7].Value = string.Empty;
                    //CODIGO VENDEDOR (VACIO)
                    dgvClientes.Rows[i].Cells[8].Value = string.Empty;
                    //CODIGO UBICACION (VACIO)
                    dgvClientes.Rows[i].Cells[9].Value = string.Empty;
                    //CODIGO CONCEPTO DE CAJA (VACIO)
                    dgvClientes.Rows[i].Cells[10].Value = string.Empty;
                    //CODIGO INSTRUMENTO FINANCIERO (VACIO)
                    dgvClientes.Rows[i].Cells[11].Value = string.Empty;
                    //CANTIDAD INSTRUMENTO FINANCIERO (VACIO)
                    dgvClientes.Rows[i].Cells[12].Value = string.Empty;
                    //CODIGO DETALLE GASTO (VACIO)
                    dgvClientes.Rows[i].Cells[13].Value = string.Empty;
                    //CANTIDAD CONCEPTO DE GASTO (VACIO)
                    dgvClientes.Rows[i].Cells[14].Value = string.Empty;
                    //CODIGO CENTRO DE COSTO (VACIO EN CUENTA DE CLIENTES)
                    dgvClientes.Rows[i].Cells[15].Value = string.Empty;
                    //TIPO DOCUMENTO CONCILIACION (VACIO)
                    dgvClientes.Rows[i].Cells[16].Value = string.Empty;
                    //NUMERO DOCUMENTO CONCILIACION
                    dgvClientes.Rows[i].Cells[17].Value = string.Empty;
                    #endregion
                    //AUXILIAR
                    string RUT = dgvLibro.Rows[i].Cells[3].Value.ToString();
                    string rutsindv = null;
                    if (RUT.Length == 10)
                    {
                        rutsindv = RUT.Substring(0, 8).Replace(".", string.Empty).Trim();
                    }
                    else if (RUT.Length == 9)
                    {
                        rutsindv = RUT.Substring(0, 7).Replace(".", string.Empty).Trim();
                    }
                    dgvClientes.Rows[i].Cells[18].Value = rutsindv;
                    //TIPO DOCUMENTO
                    if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("33"))
                    {
                        tipodcto = "A1";
                    }
                    else if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("34"))
                    {
                        tipodcto = "A2";
                    }
                    else if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("61"))
                    {
                        tipodcto = "A3";
                    }
                    dgvClientes.Rows[i].Cells[19].Value = tipodcto;
                    //NUMERO DOCUMENTO
                    dgvClientes.Rows[i].Cells[20].Value = dgvLibro.Rows[i].Cells[0].Value.ToString();
                    //FECHA EMISION
                    dgvClientes.Rows[i].Cells[21].Value = dgvLibro.Rows[i].Cells[2].Value.ToString();
                    //FECHA VENCIMIENTO
                    DateTime fec_e = Convert.ToDateTime(dgvLibro.Rows[i].Cells[2].Value.ToString());
                    DateTime fecvcto = fec_e.AddDays(5);
                    dgvClientes.Rows[i].Cells[22].Value = fecvcto.ToShortDateString();
                    //TIPO DOCUMENTO REFERENCIA
                    if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("33"))
                    {
                        tipodctoref = "A1";
                    }
                    else if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("34"))
                    {
                        tipodctoref = "A2";
                    }
                    else if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("61"))
                    {
                        tipodctoref = "A2";
                    }
                    dgvClientes.Rows[i].Cells[23].Value = tipodctoref;
                    //NUMERO DOCUMENTO REFERENCIA
                    if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("33") || dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("34"))
                    {
                        numdctoref = dgvLibro.Rows[i].Cells[0].Value.ToString();
                    }
                    else if (dgvLibro.Rows[i].Cells[1].Value.ToString().Equals("61"))
                    {
                        numdctoref = dgvLibro.Rows[i].Cells[9].Value.ToString();
                    }
                    dgvClientes.Rows[i].Cells[24].Value = numdctoref;
                    //NUMERO CORRELATIVO (VACIO)
                    dgvClientes.Rows[i].Cells[25].Value = string.Empty;
                    //MONTO 1 = MONTO EXENTO
                    dgvClientes.Rows[i].Cells[26].Value = dgvLibro.Rows[i].Cells[5].Value.ToString();
                    //MONTO 2 = MONTO AFECTO
                    dgvClientes.Rows[i].Cells[27].Value = dgvLibro.Rows[i].Cells[6].Value.ToString();
                    //MONTO 3 = IVA
                    dgvClientes.Rows[i].Cells[28].Value = dgvLibro.Rows[i].Cells[7].Value.ToString();
                    //MONTO 4
                    dgvClientes.Rows[i].Cells[29].Value = string.Empty;
                    //MONTO 5
                    dgvClientes.Rows[i].Cells[30].Value = string.Empty;
                    //MONTO 6
                    dgvClientes.Rows[i].Cells[31].Value = string.Empty;
                    //MONTO 7
                    dgvClientes.Rows[i].Cells[32].Value = string.Empty;
                    //MONTO 8
                    dgvClientes.Rows[i].Cells[33].Value = string.Empty;
                    //MONTO 9
                    dgvClientes.Rows[i].Cells[34].Value = string.Empty;
                    //MONTO SUMA
                    dgvClientes.Rows[i].Cells[35].Value = dgvLibro.Rows[i].Cells[8].Value.ToString();
                    //GRABA AL DETALLE DE LIBRO (S, PARA CUENTA CLIENTE)
                    dgvClientes.Rows[i].Cells[36].Value = "S";
                    #region vacio 2
                    //DOCUMENTO NULO (VACIO)
                    dgvClientes.Rows[i].Cells[37].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO (VACIO)
                    dgvClientes.Rows[i].Cells[38].Value = string.Empty;
                    //MONTO FLUJO 1 (VACIO)
                    dgvClientes.Rows[i].Cells[39].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 2 (VACIO)
                    dgvClientes.Rows[i].Cells[40].Value = string.Empty;
                    //MONTO FLUJO 2 (VACIO)
                    dgvClientes.Rows[i].Cells[41].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 3 (VACIO)
                    dgvClientes.Rows[i].Cells[42].Value = string.Empty;
                    //MONTO FLUJO 3 (VACIO)
                    dgvClientes.Rows[i].Cells[43].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 4 (VACIO)
                    dgvClientes.Rows[i].Cells[44].Value = string.Empty;
                    //MONTO FLUJO 4 (VACIO)
                    dgvClientes.Rows[i].Cells[45].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 5 (VACIO)
                    dgvClientes.Rows[i].Cells[46].Value = string.Empty;
                    //MONTO FLUJO 5 (VACIO)
                    dgvClientes.Rows[i].Cells[47].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 6 (VACIO)
                    dgvClientes.Rows[i].Cells[48].Value = string.Empty;
                    //MONTO FLUJO 6
                    dgvClientes.Rows[i].Cells[49].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 7 (VACIO)
                    dgvClientes.Rows[i].Cells[50].Value = string.Empty;
                    //MONTO FLUJO 7
                    dgvClientes.Rows[i].Cells[51].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 8 (VACIO)
                    dgvClientes.Rows[i].Cells[52].Value = string.Empty;
                    //MONTO FLUJO 8
                    dgvClientes.Rows[i].Cells[53].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 9 (VACIO)
                    dgvClientes.Rows[i].Cells[54].Value = string.Empty;
                    //MONTO FLUJO 9
                    dgvClientes.Rows[i].Cells[55].Value = string.Empty;
                    //CODIGO FLUJO EFECTIVO 10 (VACIO)
                    dgvClientes.Rows[i].Cells[56].Value = string.Empty;
                    //MONTO FLUJO 10 (VACIO)
                    dgvClientes.Rows[i].Cells[57].Value = string.Empty;
                    //NUMERO CUOTA PAGO (VACIO)
                    dgvClientes.Rows[i].Cells[58].Value = string.Empty;
                    //NUMERO DOCUMENTO DESDE (VACIO)
                    dgvClientes.Rows[i].Cells[59].Value = string.Empty;
                    //NUMERO DOCUMENTO HASTA (VACIO)
                    dgvClientes.Rows[i].Cells[60].Value = string.Empty;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Llenar tabla IVA
        private void TablaIVA(DateTime fecha_inicial, DateTime fecha_final)
        {
            try
            {
                dgvIVA.Rows.Clear();
                string cuenta = "2-1-03-01-01";
                string mes = dtpFinal.Value.ToString("MMMM").ToUpper();
                string year = dtpFinal.Value.Year.ToString();
                string iva = null;
                string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
                using (OdbcConnection connection = new OdbcConnection(con))
                {
                    using (OdbcCommand cmd = new OdbcCommand("select SUM(m.iva_e) as SUMA " +
                        "from movi_enc as m " +
                        "join movi_det as d on d.num_e = m.num_e and d.tipodte = m.tipodte " +
                        "where m.fec_e >= ? " +
                        "and m.fec_e <= ? " +
                        "and m.est_e <> 8 " +
                        "and m.tipodte = 33 " +
                        "GROUP BY m.tipodte ", connection))
                    {
                        cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                        cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);

                            foreach (DataRow item in dt.Rows)
                            {
                                int n = dgvIVA.Rows.Add();
                                iva = item["SUMA"].ToString();
                                dgvIVA.Rows[n].Cells[0].Value = cuenta;
                                //MONTO AL DEBE
                                dgvIVA.Rows[n].Cells[1].Value = 0;

                                //MONTO AL HABER
                                dgvIVA.Rows[n].Cells[2].Value = iva;

                                //GLOSA
                                dgvIVA.Rows[n].Cells[3].Value = "I.V.A. D/F LIBRO VENTAS " + mes + " " + year;
                                #region campos vacios 1
                                //EQUIVALENCIA (VACIO)
                                dgvIVA.Rows[n].Cells[4].Value = string.Empty;
                                //MONTO AL DEBE MONEDA ADICIONAL (VACIO)
                                dgvIVA.Rows[n].Cells[5].Value = string.Empty;
                                //MONTO AL HABER MONEDA ADICIONAL (VACIO)
                                dgvIVA.Rows[n].Cells[6].Value = string.Empty;
                                //CODIGO CONDICION VENTA (VACIO)
                                dgvIVA.Rows[n].Cells[7].Value = string.Empty;
                                //CODIGO VENDEDOR (VACIO)
                                dgvIVA.Rows[n].Cells[8].Value = string.Empty;
                                //CODIGO UBICACION (VACIO)
                                dgvIVA.Rows[n].Cells[9].Value = string.Empty;
                                //CODIGO CONCEPTO DE CAJA (VACIO)
                                dgvIVA.Rows[n].Cells[10].Value = string.Empty;
                                //CODIGO INSTRUMENTO FINANCIERO (VACIO)
                                dgvIVA.Rows[n].Cells[11].Value = string.Empty;
                                //CANTIDAD INSTRUMENTO FINANCIERO (VACIO)
                                dgvIVA.Rows[n].Cells[12].Value = string.Empty;
                                //CODIGO DETALLE GASTO (VACIO)
                                dgvIVA.Rows[n].Cells[13].Value = string.Empty;
                                //CANTIDAD CONCEPTO DE GASTO (VACIO)
                                dgvIVA.Rows[n].Cells[14].Value = string.Empty;
                                #endregion
                                //CODIGO CENTRO DE COSTO (VACIO EN CUENTA DE CLIENTES)
                                dgvIVA.Rows[n].Cells[15].Value = string.Empty;
                                #region montos vacios 2
                                //TIPO DOCUMENTO CONCILIACION (VACIO)
                                dgvIVA.Rows[n].Cells[16].Value = string.Empty;
                                //NUMERO DOCUMENTO CONCILIACION
                                dgvIVA.Rows[n].Cells[17].Value = string.Empty;
                                //AUXILIAR (VACIO)
                                dgvIVA.Rows[n].Cells[18].Value = string.Empty;
                                //TIPO DOCUMENTO (VACIO)
                                dgvIVA.Rows[n].Cells[19].Value = string.Empty;
                                //NUMERO DOCUMENTO
                                dgvIVA.Rows[n].Cells[20].Value = string.Empty;
                                //FECHA EMISION
                                dgvIVA.Rows[n].Cells[21].Value = string.Empty;
                                //FECHA VENCIMIENTO
                                dgvIVA.Rows[n].Cells[23].Value = string.Empty;
                                //NUMERO DOCUMENTO REFERENCIA
                                dgvIVA.Rows[n].Cells[24].Value = string.Empty;
                                //NUMERO CORRELATIVO (VACIO)
                                dgvIVA.Rows[n].Cells[25].Value = string.Empty;
                                //MONTO 1 = MONTO EXENTO
                                dgvIVA.Rows[n].Cells[26].Value = string.Empty;
                                //MONTO 2 = MONTO AFECTO
                                dgvIVA.Rows[n].Cells[27].Value = string.Empty;
                                //MONTO 3 = IVA
                                dgvIVA.Rows[n].Cells[28].Value = string.Empty;
                                //MONTO 4
                                dgvIVA.Rows[n].Cells[29].Value = string.Empty;
                                //MONTO 5
                                dgvIVA.Rows[n].Cells[30].Value = string.Empty;
                                //MONTO 6
                                dgvIVA.Rows[n].Cells[31].Value = string.Empty;
                                //MONTO 7
                                dgvIVA.Rows[n].Cells[32].Value = string.Empty;
                                //MONTO 8
                                dgvIVA.Rows[n].Cells[33].Value = string.Empty;
                                //MONTO 9
                                dgvIVA.Rows[n].Cells[34].Value = string.Empty;
                                //MONTO SUMA
                                dgvIVA.Rows[n].Cells[35].Value = string.Empty;
                                #endregion
                                //GRABA AL DETALLE DE LIBRO (S, PARA CUENTA CLIENTE)
                                dgvIVA.Rows[n].Cells[36].Value = "N";
                                #region campos vacios 3
                                //DOCUMENTO NULO (VACIO)
                                dgvIVA.Rows[n].Cells[37].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO (VACIO)
                                dgvIVA.Rows[n].Cells[38].Value = string.Empty;
                                //MONTO FLUJO 1 (VACIO)
                                dgvIVA.Rows[n].Cells[39].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 2 (VACIO)
                                dgvIVA.Rows[n].Cells[40].Value = string.Empty;
                                //MONTO FLUJO 2 (VACIO)
                                dgvIVA.Rows[n].Cells[41].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 3 (VACIO)
                                dgvIVA.Rows[n].Cells[42].Value = string.Empty;
                                //MONTO FLUJO 3 (VACIO)
                                dgvIVA.Rows[n].Cells[43].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 4 (VACIO)
                                dgvIVA.Rows[n].Cells[44].Value = string.Empty;
                                //MONTO FLUJO 4 (VACIO)
                                dgvIVA.Rows[n].Cells[45].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 5 (VACIO)
                                dgvIVA.Rows[n].Cells[46].Value = string.Empty;
                                //MONTO FLUJO 5 (VACIO)
                                dgvIVA.Rows[n].Cells[47].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 6 (VACIO)
                                dgvIVA.Rows[n].Cells[48].Value = string.Empty;
                                //MONTO FLUJO 6
                                dgvIVA.Rows[n].Cells[49].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 7 (VACIO)
                                dgvIVA.Rows[n].Cells[50].Value = string.Empty;
                                //MONTO FLUJO 7
                                dgvIVA.Rows[n].Cells[51].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 8 (VACIO)
                                dgvIVA.Rows[n].Cells[52].Value = string.Empty;
                                //MONTO FLUJO 8
                                dgvIVA.Rows[n].Cells[53].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 9 (VACIO)
                                dgvIVA.Rows[n].Cells[54].Value = string.Empty;
                                //MONTO FLUJO 9
                                dgvIVA.Rows[n].Cells[55].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 10 (VACIO)
                                dgvIVA.Rows[n].Cells[56].Value = string.Empty;
                                //MONTO FLUJO 10 (VACIO)
                                dgvIVA.Rows[n].Cells[57].Value = string.Empty;
                                //NUMERO CUOTA PAGO (VACIO)
                                dgvIVA.Rows[n].Cells[58].Value = string.Empty;
                                //NUMERO DOCUMENTO DESDE (VACIO)
                                dgvIVA.Rows[n].Cells[59].Value = string.Empty;
                                //NUMERO DOCUMENTO HASTA (VACIO)
                                dgvIVA.Rows[n].Cells[60].Value = string.Empty;
                                #endregion
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Llenar tabla códigos de ingreso
        private void TablaIngreso(DateTime fecha_inicial, DateTime fecha_final) {
            try
            {
                dgvIngresos.Rows.Clear();
                string mes = dtpFinal.Value.ToString("MMMM").ToUpper();
                string year = dtpFinal.Value.Year.ToString();
                string text = null;
                string dte = null;
                string con = ConfigurationManager.ConnectionStrings["SR"].ConnectionString;
                using (OdbcConnection connection = new OdbcConnection(con))
                {
                    using (OdbcCommand cmd = new OdbcCommand("select distinct(a.cta_softland) as CTA, SUM(d.p_real) as SUMA, m.tipodte as DTE " +
                        "from movi_enc as m " +
                        "join clientes as c on c.rut = m.rut_e " +
                        "join movi_det as d on d.num_e = m.num_e and d.tipodte = m.tipodte " +
                        "join articulo as a on a.codigo = d.cod_d " +
                        "where m.fec_e >= ? " +
                        "and m.fec_e <= ? " +
                        "and m.est_e <> 8 " +
                        "and m.tipodte IN (33, 34, 61) " +
                        "GROUP BY a.cta_softland, m.tipodte " +
                        "order by m.tipodte ", connection))
                    {
                        cmd.Parameters.Add(new OdbcParameter("@fi", fecha_inicial));
                        cmd.Parameters.Add(new OdbcParameter("@ff", fecha_final));
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);

                            foreach (DataRow item in dt.Rows)
                            {
                                int n = dgvIngresos.Rows.Add();
                                dte = item["DTE"].ToString();
                                dgvIngresos.Rows[n].Cells[0].Value = item["CTA"].ToString();
                                //ORDENA MONTOS A LA CUENTA DEBE O HABER DEPENDIENDO DEL TIPO DE DOCUMENTO
                                if (dte.Equals("33") || dte.Equals("34"))
                                {
                                    //MONTO AL HABER
                                    dgvIngresos.Rows[n].Cells[2].Value = item["SUMA"].ToString();
                                }
                                else
                                {
                                    //MONTO AL DEBE
                                    dgvIngresos.Rows[n].Cells[1].Value = item["SUMA"].ToString();
                                }
                                //GLOSA
                                if (dte.Equals("33"))
                                {
                                    text = "INGRESOS POR VENTAS FACTURAS AFECTA MES " + mes + " " + year;
                                }
                                else if (dte.Equals("34"))
                                {
                                    text = "INGRESOS POR VENTAS FACTURAS EXENTAS MES " + mes + " " + year;
                                }
                                else if (dte.Equals("61"))
                                {
                                    text = "AJUSTE NOTAS DE CREDITO MES " + mes + " " + year;
                                }

                                dgvIngresos.Rows[n].Cells[3].Value = text;
                                #region campos vacios 1
                                //EQUIVALENCIA (VACIO)
                                dgvIngresos.Rows[n].Cells[4].Value = string.Empty;
                                //MONTO AL DEBE MONEDA ADICIONAL (VACIO)
                                dgvIngresos.Rows[n].Cells[5].Value = string.Empty;
                                //MONTO AL HABER MONEDA ADICIONAL (VACIO)
                                dgvIngresos.Rows[n].Cells[6].Value = string.Empty;
                                //CODIGO CONDICION VENTA (VACIO)
                                dgvIngresos.Rows[n].Cells[7].Value = string.Empty;
                                //CODIGO VENDEDOR (VACIO)
                                dgvIngresos.Rows[n].Cells[8].Value = string.Empty;
                                //CODIGO UBICACION (VACIO)
                                dgvIngresos.Rows[n].Cells[9].Value = string.Empty;
                                //CODIGO CONCEPTO DE CAJA (VACIO)
                                dgvIngresos.Rows[n].Cells[10].Value = string.Empty;
                                //CODIGO INSTRUMENTO FINANCIERO (VACIO)
                                dgvIngresos.Rows[n].Cells[11].Value = string.Empty;
                                //CANTIDAD INSTRUMENTO FINANCIERO (VACIO)
                                dgvIngresos.Rows[n].Cells[12].Value = string.Empty;
                                //CODIGO DETALLE GASTO (VACIO)
                                dgvIngresos.Rows[n].Cells[13].Value = string.Empty;
                                //CANTIDAD CONCEPTO DE GASTO (VACIO)
                                dgvIngresos.Rows[n].Cells[14].Value = string.Empty;
                                #endregion
                                //CODIGO CENTRO DE COSTO (VACIO EN CUENTA DE CLIENTES)
                                dgvIngresos.Rows[n].Cells[15].Value = "02-007";
                                #region montos vacios 2
                                //TIPO DOCUMENTO CONCILIACION (VACIO)
                                dgvIngresos.Rows[n].Cells[16].Value = string.Empty;
                                //NUMERO DOCUMENTO CONCILIACION
                                dgvIngresos.Rows[n].Cells[17].Value = string.Empty;
                                //AUXILIAR (VACIO)
                                dgvIngresos.Rows[n].Cells[18].Value = string.Empty;
                                //TIPO DOCUMENTO (VACIO)
                                dgvIngresos.Rows[n].Cells[19].Value = string.Empty;
                                //NUMERO DOCUMENTO
                                dgvIngresos.Rows[n].Cells[20].Value = string.Empty;
                                //FECHA EMISION
                                dgvIngresos.Rows[n].Cells[21].Value = string.Empty;
                                //FECHA VENCIMIENTO
                                dgvIngresos.Rows[n].Cells[23].Value = string.Empty;
                                //NUMERO DOCUMENTO REFERENCIA
                                dgvIngresos.Rows[n].Cells[24].Value = string.Empty;
                                //NUMERO CORRELATIVO (VACIO)
                                dgvIngresos.Rows[n].Cells[25].Value = string.Empty;
                                //MONTO 1 = MONTO EXENTO
                                dgvIngresos.Rows[n].Cells[26].Value = string.Empty;
                                //MONTO 2 = MONTO AFECTO
                                dgvIngresos.Rows[n].Cells[27].Value = string.Empty;
                                //MONTO 3 = IVA
                                dgvIngresos.Rows[n].Cells[28].Value = string.Empty;
                                //MONTO 4
                                dgvIngresos.Rows[n].Cells[29].Value = string.Empty;
                                //MONTO 5
                                dgvIngresos.Rows[n].Cells[30].Value = string.Empty;
                                //MONTO 6
                                dgvIngresos.Rows[n].Cells[31].Value = string.Empty;
                                //MONTO 7
                                dgvIngresos.Rows[n].Cells[32].Value = string.Empty;
                                //MONTO 8
                                dgvIngresos.Rows[n].Cells[33].Value = string.Empty;
                                //MONTO 9
                                dgvIngresos.Rows[n].Cells[34].Value = string.Empty;
                                //MONTO SUMA
                                dgvIngresos.Rows[n].Cells[35].Value = string.Empty;
                                #endregion
                                //GRABA AL DETALLE DE LIBRO (S, PARA CUENTA CLIENTE)
                                dgvIngresos.Rows[n].Cells[36].Value = "N";
                                #region campos vacios 3
                                //DOCUMENTO NULO (VACIO)
                                dgvIngresos.Rows[n].Cells[37].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO (VACIO)
                                dgvIngresos.Rows[n].Cells[38].Value = string.Empty;
                                //MONTO FLUJO 1 (VACIO)
                                dgvIngresos.Rows[n].Cells[39].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 2 (VACIO)
                                dgvIngresos.Rows[n].Cells[40].Value = string.Empty;
                                //MONTO FLUJO 2 (VACIO)
                                dgvIngresos.Rows[n].Cells[41].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 3 (VACIO)
                                dgvIngresos.Rows[n].Cells[42].Value = string.Empty;
                                //MONTO FLUJO 3 (VACIO)
                                dgvIngresos.Rows[n].Cells[43].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 4 (VACIO)
                                dgvIngresos.Rows[n].Cells[44].Value = string.Empty;
                                //MONTO FLUJO 4 (VACIO)
                                dgvIngresos.Rows[n].Cells[45].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 5 (VACIO)
                                dgvIngresos.Rows[n].Cells[46].Value = string.Empty;
                                //MONTO FLUJO 5 (VACIO)
                                dgvIngresos.Rows[n].Cells[47].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 6 (VACIO)
                                dgvIngresos.Rows[n].Cells[48].Value = string.Empty;
                                //MONTO FLUJO 6
                                dgvIngresos.Rows[n].Cells[49].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 7 (VACIO)
                                dgvIngresos.Rows[n].Cells[50].Value = string.Empty;
                                //MONTO FLUJO 7
                                dgvIngresos.Rows[n].Cells[51].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 8 (VACIO)
                                dgvIngresos.Rows[n].Cells[52].Value = string.Empty;
                                //MONTO FLUJO 8
                                dgvIngresos.Rows[n].Cells[53].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 9 (VACIO)
                                dgvIngresos.Rows[n].Cells[54].Value = string.Empty;
                                //MONTO FLUJO 9
                                dgvIngresos.Rows[n].Cells[55].Value = string.Empty;
                                //CODIGO FLUJO EFECTIVO 10 (VACIO)
                                dgvIngresos.Rows[n].Cells[56].Value = string.Empty;
                                //MONTO FLUJO 10 (VACIO)
                                dgvIngresos.Rows[n].Cells[57].Value = string.Empty;
                                //NUMERO CUOTA PAGO (VACIO)
                                dgvIngresos.Rows[n].Cells[58].Value = string.Empty;
                                //NUMERO DOCUMENTO DESDE (VACIO)
                                dgvIngresos.Rows[n].Cells[59].Value = string.Empty;
                                //NUMERO DOCUMENTO HASTA (VACIO)
                                dgvIngresos.Rows[n].Cells[60].Value = string.Empty;
                                #endregion
                            }
                        }
                    }
                }
                SumaDebe();
                SumaHaber();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region SUMA COLUMNAS DEBE DE LAS 3 TABLAS
        private void SumaDebe()
        {
            int totalClientes = dgvClientes.Rows.Cast<DataGridViewRow>()
                .Sum(t => Convert.ToInt32(t.Cells[1].Value));
            int totalIva = 0;
            for (int i = 0; i < dgvIVA.Rows.Count; i++)
            {
                if (dgvIVA.Rows[i].Cells[2].Value.ToString() != null)
                {
                    totalIva = dgvIVA.Rows.Cast<DataGridViewRow>()
               .Sum(t => Convert.ToInt32(t.Cells[1].Value));
                }
                else
                {
                    totalIva = 0;
                }
            }
            int totalIngresos = dgvIngresos.Rows.Cast<DataGridViewRow>()
                .Sum(t => Convert.ToInt32(t.Cells[1].Value));
            int suma = totalClientes + totalIva + totalIngresos;
            txtDebe.Text = suma.ToString("C");
        }
        #endregion
        #region SUMA COLUMNAS HABER DE LAS 3 TABLAS
        private void SumaHaber()
        {
            int totalClientes = dgvClientes.Rows.Cast<DataGridViewRow>()
                .Sum(t => Convert.ToInt32(t.Cells[2].Value));
            int totalIva = 0;
            for (int i = 0; i < dgvIVA.Rows.Count; i++)
            {
                if (dgvIVA.Rows[i].Cells[2].Value.ToString() != null)
                {
                    totalIva = dgvIVA.Rows.Cast<DataGridViewRow>()
               .Sum(t => Convert.ToInt32(t.Cells[2].Value));
                }
                else
                {
                    totalIva = 0;
                }
            }
            int totalIngresos = dgvIngresos.Rows.Cast<DataGridViewRow>()
                .Sum(t => Convert.ToInt32(t.Cells[2].Value));
            int suma = totalClientes + totalIva + totalIngresos;
            txtHaber.Text = suma.ToString("C");
        }
        #endregion
        #region Exportar datos a csv
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "CSV|*.csv", ValidateNames = true })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        using (StreamWriter sw = new StreamWriter(new FileStream(sfd.FileName, FileMode.Create), new UTF8Encoding(false)))
                        {
                            var sb = new StringBuilder();

                            foreach (DataGridViewRow row in dgvClientes.Rows)
                            {
                                var cells = row.Cells.Cast<DataGridViewCell>();
                                sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value + "\"").ToArray()));
                            }

                            foreach (DataGridViewRow row in dgvIVA.Rows)
                            {
                                var cells = row.Cells.Cast<DataGridViewCell>();
                                sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value + "\"").ToArray()));
                            }

                            foreach (DataGridViewRow row in dgvIngresos.Rows)
                            {
                                var cells = row.Cells.Cast<DataGridViewCell>();
                                sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value + "\"").ToArray()));
                            }

                            sw.Write(sb.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region Inicialización de notificación
        public void notificacion(string mensaje)
        {
            notificacionInicio.BalloonTipIcon = ToolTipIcon.Info;
            notificacionInicio.BalloonTipTitle = "Actualización Disponible";
            notificacionInicio.BalloonTipText = mensaje;
            notificacionInicio.ShowBalloonTip(5000);
        }
        #endregion
        #region Comprobar Actualizaciones
        private void comprobarActualizacionesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CheckForUpdate();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        #endregion
        #region Cerrar programa desde ícono de notificación
        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        #endregion

    }
}
