using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Respuesta;

namespace TRADINGVSTO
{
    public partial class ThisAddIn
    {
        public class Objeto
        {
            public String Nombre { get; set; }

        }
        Excel.Worksheet _sheet1;
        private List<Objeto> _objeto;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _objeto = new List<Objeto> {
            new Objeto { Nombre = "ResumenMensual"},
            new Objeto { Nombre= "ResumenAnual" }
            };
            Application.WorkbookActivate +=
           Application_ActiveWorkbookChanges;
            Application.WorkbookDeactivate += Application_ActiveWorkbookChanges;
            Globals.ThisAddIn.Application.SheetSelectionChange += activeSheet_SelectionChange;
            Application.SheetBeforeDoubleClick += Application_SheetBeforeDoubleClick;
            _sheet1 = (Excel.Worksheet)Application.ActiveSheet;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            Startup += (ThisAddIn_Startup);
            Shutdown += (ThisAddIn_Shutdown);
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #endregion

        void Application_ActiveWorkbookChanges(Excel.Workbook wb)
        {
            //TODO: Active Workbook has changed. Ribbon should be updated.    
            //wb.Unprotect();
        }
        void activeSheet_SelectionChange(object sh, Excel.Range target)
        {
            _sheet1 = (Excel.Worksheet)sh;
            if (target.Row != 1 && (_objeto.FirstOrDefault(x => x.Nombre == _sheet1.Name) != null))
            {
                try
                {
                    _sheet1.Unprotect();
                    Globals.ThisAddIn.Application.Cells.Locked = false;
                    //BloquearRango(_rowCount);
                    _sheet1.Protect(AllowSorting: true, AllowFiltering: true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            else
            {
                _sheet1.Unprotect();
            }
        }

        private Excel.Worksheet _reporte;
        public Excel.Worksheet InicializarExcelConTemplate(string nombreHoja)
        {
            try
            {
                _sheet1 = (Excel.Worksheet)Application.ActiveSheet;
                _sheet1.Unprotect();
                var found = Application.Sheets.Cast<Excel.Worksheet>().Any(sheet => sheet.Name == nombreHoja);
                var awa = Application.Workbooks.Application;//nueva app
                if (!found)
                {
                    var ows = Application.Worksheets[1];// excel actual
                    var sPath = Path.GetTempFileName(); // archivo temporal
                    File.WriteAllBytes(sPath, Properties.Resources.TABLATRADING);//se copia el template
                    var oTemplate = Application.Workbooks.Add(sPath); //path del template temporal  
                    var worksheet = oTemplate.Worksheets[nombreHoja] as Excel.Worksheet;//descripcion del template
                    worksheet?.Copy(After: ows); oTemplate.Close(false, missing, missing);
                    File.Delete(sPath);
                }
                _reporte = awa.Worksheets[nombreHoja] as Excel.Worksheet;//descripcion de la hoja actual   
                _reporte?.Activate();
            }
            catch (Exception e)
            {
                  MessageBox.Show(e.Message);
            }
            return _reporte;
        }
        public static DateTime FirstDayOfWeek(DateTime date)
        {
            var fdow = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
            var offset = fdow - date.DayOfWeek;
            var fdowDate = date.AddDays(offset);
            return fdowDate;
        }
        public static DateTime LastDayOfWeek(DateTime date)
        {
            var ldowDate = FirstDayOfWeek(date).AddDays(6);
            return ldowDate;
        }

        public string Lunesdelmes;
        public int Contador;
        public List<DatosSemanal> Trade;
        public void ResumenSemanal(List<DatosSemanal> listaSemana)
        {   
            Application.ScreenUpdating = false;
            var excelazo = InicializarExcelConTemplate("ResumenMensual");
            string name = Globals.ThisAddIn.Application.ActiveSheet.Name;
            //if (name.Equals(@"ResumenMensual")==ib)
            //{
            //    //excelazo.Range["B" + 7 + ":M" + Globals.ThisAddIn.Application.ActiveSheet.Cells.Find("*", Missing.Value,
            ////    Missing.Value, Missing.Value, Excel.XlSearchOrder.xlByRows,
            ////  Excel.XlSearchDirection.xlPrevious, false, Missing.Value,
            ////   Missing.Value).Row + 1].Value2 = "";
            //    excelazo = InicializarExcelConTemplate("ResumenMensual");
            //}
            var nueva = new List<string>();
            for (var i = 0; i <listaSemana.Count; i++)
            {
                var diainicio = FirstDayOfWeek(Convert.ToDateTime(listaSemana[i].Fecha)).ToShortDateString();
                var value =diainicio;
                const char delimite = '/';
                var substrings = value.Split(delimite);
                Lunesdelmes= substrings[0];
                nueva.Add(Lunesdelmes);
            }
            //var grouped = nueva
            //       .GroupBy(s => s).Select(g => new { Symbol = g.Key});
            //var lunes = grouped.Select(a => a.Symbol).ToArray();  
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var x = 0; x < listaSemana.Count; x++)
            {
                var fechaactual = DateTime.Now;
                var iniciomes = new DateTime(fechaactual.Year, fechaactual.Month, 1).ToShortDateString();
                var finmes = new DateTime(fechaactual.Year, fechaactual.Month + 1, 1).AddDays(-1).ToString("yyyy-MM-dd");
              
                var lunesprimero = FirstDayOfWeek(Convert.ToDateTime(iniciomes)).ToString("yyyy-MM-dd");
                var domingoprimero = LastDayOfWeek(Convert.ToDateTime(iniciomes));
                var viernesprimero = domingoprimero.AddDays(-2).ToString("yyyy-MM-dd");
                var lunessegundo = FirstDayOfWeek(Convert.ToDateTime(domingoprimero).AddDays(1)).ToString("yyyy-MM-dd");
                var domingosegundo = LastDayOfWeek(Convert.ToDateTime(domingoprimero));
                var viernessegundo = domingosegundo.AddDays(-2).ToString("yyyy-MM-dd");
                var lunestercero = FirstDayOfWeek(Convert.ToDateTime(domingosegundo).AddDays(1)).ToString("yyyy-MM-dd");
                var domingotercero = LastDayOfWeek(Convert.ToDateTime(domingosegundo));
                var viernestercero = domingotercero.AddDays(-2).ToString("yyyy-MM-dd");
                var lunescuarto = FirstDayOfWeek(Convert.ToDateTime(domingotercero).AddDays(1)).ToString("yyyy-MM-dd");
                var domingocuarto = LastDayOfWeek(Convert.ToDateTime(domingotercero));
                var viernescuarto = domingocuarto.AddDays(-2).ToString("yyyy-MM-dd");
                var lunesultimo = FirstDayOfWeek(Convert.ToDateTime(finmes)).ToString("yyyy-MM-dd");
                //var domingoultimo = LastDayOfWeek(Convert.ToDateTime(finmes)).ToString("yyyy-MM-dd");
                var dia = listaSemana[x].Fecha;
                var dato = dia;
                const char split = '-';
                var dialista = dato.Split(split);
                var diafor = Convert.ToInt32(dialista[2]);
                const char split1 = '-';
                var ultimo = viernesprimero.Split(split1);
                var ultimodia = Convert.ToInt32(ultimo[2]);
                var ultimo2 = viernessegundo.Split(split1);
                var ultimodia2 = Convert.ToInt32(ultimo2[2]);
                var ultimo3 = viernestercero.Split(split1);
                var ultimodia3 = Convert.ToInt32(ultimo3[2]);
                var ultimo4 = viernescuarto.Split(split1);
                var ultimodia4 = Convert.ToInt32(ultimo4[2]);
                var ultimo5 = lunesultimo.Split(split1);
                var ultimodia5 = Convert.ToInt32(ultimo5[2]);

                if (Convert.ToDateTime(listaSemana[x].Fecha) >= Convert.ToDateTime(lunesprimero) && Convert.ToDateTime(listaSemana[x].Fecha) <= Convert.ToDateTime(viernesprimero)
                   )
                {
                    const int semana1 = 7;
                    if (diafor == ultimodia-4)
                    {
                        Contador = semana1;
                    }
                    if (diafor == ultimodia - 3)
                    {
                        Contador = semana1 + 1;
                    }
                    if (diafor == ultimodia - 2)
                    {
                        Contador = semana1 + 2;
                    }
                    if (diafor == ultimodia - 1)
                    {
                        Contador = semana1 + 3;
                    }
                    if (diafor == ultimodia)
                    {
                        Contador = semana1 + 4;
                    }
                    if (excelazo != null)
                    {
                        (excelazo.Range["B" + Contador]).Value2 = listaSemana[x].Fecha;
                        (excelazo.Range["C" + Contador]).Value2 = listaSemana[x].TipoTrade;
                        (excelazo.Range["D" + Contador]).Value2 = listaSemana[x].Proceso;
                        (excelazo.Range["F" + Contador]).Value2 = listaSemana[x].Direccion;
                        (excelazo.Range["G" + Contador]).Value2 = listaSemana[x].Puntos;
                        (excelazo.Range["K" + Contador]).Value2 = listaSemana[x].Exito;
                        (excelazo.Range["L" + Contador]).Value2 = listaSemana[x].Fracaso;
                        (excelazo.Range["M" + Contador]).Value2 = listaSemana[x].Id;
                    }
                }
                if (Convert.ToDateTime(listaSemana[x].Fecha) >= Convert.ToDateTime(lunessegundo) && Convert.ToDateTime(listaSemana[x].Fecha) <= Convert.ToDateTime(viernessegundo))
                {
                    const int semana2 = 16;
                    if (diafor == ultimodia2-4)
                    {
                        Contador = semana2;
                    }
                    if (diafor == ultimodia2 - 3)
                    {
                        Contador = semana2 + 1;
                    }
                    if (diafor == ultimodia2 - 2)
                    {
                        Contador = semana2 + 2;
                    }
                    if (diafor == ultimodia2 - 1)
                    {
                        Contador = semana2 + 3;
                    }
                    if (diafor == ultimodia2)
                    {
                        Contador = semana2 + 4;
                    }
                    if (excelazo != null)
                    {
                        (excelazo.Range["B" + Contador]).Value2 = listaSemana[x].Fecha;
                        (excelazo.Range["C" + Contador]).Value2 = listaSemana[x].TipoTrade;
                        (excelazo.Range["D" + Contador]).Value2 = listaSemana[x].Proceso;
                        (excelazo.Range["F" + Contador]).Value2 = listaSemana[x].Direccion;
                        (excelazo.Range["G" + Contador]).Value2 = listaSemana[x].Puntos;
                        (excelazo.Range["K" + Contador]).Value2 = listaSemana[x].Exito;
                        (excelazo.Range["L" + Contador]).Value2 = listaSemana[x].Fracaso;
                        (excelazo.Range["M" + Contador]).Value2 = listaSemana[x].Id;
                    }
                }
                if (Convert.ToDateTime(listaSemana[x].Fecha) >= Convert.ToDateTime(lunestercero) && Convert.ToDateTime(listaSemana[x].Fecha) <= Convert.ToDateTime(viernestercero))
                {
                    const int semana3 = 25;
                    if (diafor == ultimodia3 - 4)
                    {
                        Contador = semana3;
                    }
                    if (diafor == ultimodia3 - 3)
                    {
                        Contador = semana3 + 1;
                    }
                    if (diafor == ultimodia3 - 2)
                    {
                        Contador = semana3 + 2;
                    }
                    if (diafor == ultimodia3 - 1)
                    {
                        Contador = semana3 + 3;
                    }
                    if (diafor == ultimodia3)
                    {
                        Contador = semana3 + 4;
                    }
                    if (excelazo != null)
                    {
                        (excelazo.Range["B" + Contador]).Value2 = listaSemana[x].Fecha;
                        (excelazo.Range["C" + Contador]).Value2 = listaSemana[x].TipoTrade;
                        (excelazo.Range["D" + Contador]).Value2 = listaSemana[x].Proceso;
                        (excelazo.Range["F" + Contador]).Value2 = listaSemana[x].Direccion;
                        (excelazo.Range["G" + Contador]).Value2 = listaSemana[x].Puntos;
                        (excelazo.Range["K" + Contador]).Value2 = listaSemana[x].Exito;
                        (excelazo.Range["L" + Contador]).Value2 = listaSemana[x].Fracaso;
                        (excelazo.Range["M" + Contador]).Value2 = listaSemana[x].Id;
                    }
                }
                if (Convert.ToDateTime(listaSemana[x].Fecha) >= Convert.ToDateTime(lunescuarto) && Convert.ToDateTime(listaSemana[x].Fecha) <= Convert.ToDateTime(viernescuarto)
                   /*Convert.ToInt32(diafor) >= Convert.ToInt32(lunes4) && Convert.ToInt32(diafor) <= Convert.ToInt32(lunes4 + 4)*/)
                {
                    const int semana4 = 34;
                    if (diafor == ultimodia4 - 4)
                    {
                        Contador = semana4;
                    }
                    if (diafor == ultimodia4 - 3)
                    {
                        Contador = semana4 + 1;
                    }
                    if (diafor == ultimodia4 - 2)
                    {
                        Contador = semana4 + 2;
                    }
                    if (diafor == ultimodia4 - 1)
                    {
                        Contador = semana4 + 3;
                    }
                    if (diafor == ultimodia4)
                    {
                        Contador = semana4 + 4;
                    }
                    if (excelazo != null)
                    {
                        (excelazo.Range["B" + Contador]).Value2 = listaSemana[x].Fecha;
                        (excelazo.Range["C" + Contador]).Value2 = listaSemana[x].TipoTrade;
                        (excelazo.Range["D" + Contador]).Value2 = listaSemana[x].Proceso;
                        (excelazo.Range["F" + Contador]).Value2 = listaSemana[x].Direccion;
                        (excelazo.Range["G" + Contador]).Value2 = listaSemana[x].Puntos;
                        (excelazo.Range["K" + Contador]).Value2 = listaSemana[x].Exito;
                        (excelazo.Range["L" + Contador]).Value2 = listaSemana[x].Fracaso;
                        (excelazo.Range["M" + Contador]).Value2 = listaSemana[x].Id;
                    }
                }
                if (Convert.ToDateTime(listaSemana[x].Fecha) >=Convert.ToDateTime(lunesultimo) && Convert.ToDateTime(listaSemana[x].Fecha) <= Convert.ToDateTime(finmes)
                    /*Convert.ToInt32(diafor) >= Convert.ToInt32(lunes5) && Convert.ToInt32(diafor) <= Convert.ToInt32(lunes5 + 4*/)
                {
                    const int semana5 = 43;
                    if (diafor == ultimodia5)
                    {
                        Contador = semana5;
                    }
                    if (diafor == ultimodia5 + 1)
                    {
                        Contador = semana5 + 1;
                    }
                    if (diafor == ultimodia5 + 2)
                    {
                        Contador = semana5 + 2;
                    }
                    if (diafor == ultimodia5 + 3)
                    {
                        Contador = semana5 + 3;
                    }
                    if (diafor == ultimodia5 + 4)
                    {
                        Contador = semana5 + 4;
                    }
                    if (excelazo != null)
                    {
                        (excelazo.Range["B" + Contador]).Value2 = listaSemana[x].Fecha;
                        (excelazo.Range["C" + Contador]).Value2 = listaSemana[x].TipoTrade;
                        (excelazo.Range["D" + Contador]).Value2 = listaSemana[x].Proceso;
                        (excelazo.Range["F" + Contador]).Value2 = listaSemana[x].Direccion;
                        (excelazo.Range["G" + Contador]).Value2 = listaSemana[x].Puntos;
                        (excelazo.Range["K" + Contador]).Value2 = listaSemana[x].Exito;
                        (excelazo.Range["L" + Contador]).Value2 = listaSemana[x].Fracaso;
                        (excelazo.Range["M" + Contador]).Value2 = listaSemana[x].Id;
                    }
                }
            }
            Application.Cells.Locked = false;
            Application.ScreenUpdating = true;
        }

        // ReSharper disable once FunctionComplexityOverflow
        public void ReporteAnual(List<DatosAnual> listaAnual,List<DatosAnual> listaGanador,List<DatosAnual> listaPerdedor,
                                  List<DatosAnual> listaTrades,List<DatosAnual> listaExito,List<DatosAnual> listaFracaso)
        {
            Application.ScreenUpdating = false;
            var excelazo = InicializarExcelConTemplate("ResumenAnual");
            //excelazo.Range["C" + 8 + Globals.ThisAddIn.Application.ActiveSheet.Cells.Find("H"+ 19, Missing.Value,
            //           Missing.Value, Missing.Value, Excel.XlSearchOrder.xlByRows,
            //          Excel.XlSearchDirection.xlPrevious, false, Missing.Value,
            //           Missing.Value).Row + 1].Value2 = "";
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < listaAnual.Count; i++)
            {
                var diafor = listaAnual[i].Mes;
                const int row= 8;
                if (diafor == "Enero")
                {
                    Contador = row;
                }
                if (diafor == "Febrero")
                {
                    Contador = row+1;
                }
                if (diafor == "Marzo")
                {
                    Contador = row+2;
                }
                if (diafor == "Abril")
                {
                    Contador = row+3;
                }
                if (diafor == "Mayo")
                {
                    Contador = row+4;
                }
                if (diafor == "Junio")
                {
                    Contador = row+5;
                }
                if (diafor == "Julio")
                {
                    Contador = row+6;
                }
                if (diafor == "Agosto")
                {
                    Contador = row+7;
                }
                if (diafor == "Septiembre")
                {
                    Contador = row+8;
                }
                if (diafor == "Octubre")
                {
                    Contador = row+9;
                }
                if (diafor == "Noviembre")
                {
                    Contador = row+10;
                }
                if (diafor == "Diciemnre")
                {
                    Contador = row + 11;
                }
                (excelazo.Range["C" + Contador ]).Value2 = listaAnual[i].Trades;
                (excelazo.Range["D" + Contador]).Value2 = listaAnual[i].Win;
                (excelazo.Range["E" + Contador]).Value2 = listaAnual[i].Loss;
                (excelazo.Range["F" + Contador]).Value2 = listaAnual[i].PuntosWin;
                (excelazo.Range["G" + Contador]).Value2 = listaAnual[i].PuntosLoss;
                (excelazo.Range["H" + Contador]).Value2 = listaAnual[i].PuntosTotales;
                (excelazo.Range["N" + Contador]).Value2 = listaAnual[i].Largo;
                (excelazo.Range["O" + Contador]).Value2 = listaAnual[i].Corto;
            }
            var enerowin = listaGanador.Where(y => y.Mes =="Enero").ToList().Max(y=> y.TradeWin);
            var febrerowin= listaGanador.Where(y => y.Mes == "Febrero").ToList().Max(y => y.TradeWin);
            var marzowin= listaGanador.Where(y => y.Mes == "Marzo").ToList().Max(y => y.TradeWin);
            var abrilwin = listaGanador.Where(y => y.Mes == "Abril").ToList().Max(y => y.TradeWin);
            var mayowin = listaGanador.Where(y => y.Mes == "Mayo").ToList().Max(y => y.TradeWin);
            var juniowin = listaGanador.Where(y => y.Mes == "Junio").ToList().Max(y => y.TradeWin);
            var juliowin = listaGanador.Where(y => y.Mes == "Julio").ToList().Max(y => y.TradeWin);
            var agostowin = listaGanador.Where(y => y.Mes == "Agosto").ToList().Max(y => y.TradeWin);
            var septiembrewin = listaGanador.Where(y => y.Mes == "Septiembre").ToList().Max(y => y.TradeWin);
            var octubrewin = listaGanador.Where(y => y.Mes == "Octubre").ToList().Max(y => y.TradeWin);
            var noviembrewin = listaGanador.Where(y => y.Mes == "Noviembre").ToList().Max(y => y.TradeWin);
            var diciembrewin = listaGanador.Where(y => y.Mes == "Diciembre").ToList().Max(y => y.TradeWin);
            var eneroloss = listaPerdedor.Where(y => y.Mes == "Enero").ToList().Max(y => y.TradeLoss);
            var febreroloss = listaPerdedor.Where(y => y.Mes == "Febrero").ToList().Max(y => y.TradeLoss);
            var marzoloss = listaPerdedor.Where(y => y.Mes == "Marzo").ToList().Max(y => y.TradeLoss);
            var abrilloss = listaPerdedor.Where(y => y.Mes == "Abril").ToList().Max(y => y.TradeLoss);
            var mayoloss = listaPerdedor.Where(y => y.Mes == "Mayo").ToList().Max(y => y.TradeLoss);
            var junioloss = listaPerdedor.Where(y => y.Mes == "Junio").ToList().Max(y => y.TradeLoss);
            var julioloss = listaPerdedor.Where(y => y.Mes == "Julio").ToList().Max(y => y.TradeLoss);
            var agostoloss= listaPerdedor.Where(y => y.Mes == "Agosto").ToList().Max(y => y.TradeLoss);
            var septiembreloss = listaPerdedor.Where(y => y.Mes == "Septiembre").ToList().Max(y => y.TradeLoss);
            var octubreloss = listaPerdedor.Where(y => y.Mes == "Octubre").ToList().Max(y => y.TradeLoss);
            var noviembreloss = listaPerdedor.Where(y => y.Mes == "Noviembre").ToList().Max(y => y.TradeLoss);
            var diciembreloss = listaPerdedor.Where(y => y.Mes == "Diciembre").ToList().Max(y => y.TradeLoss);
            (excelazo.Range["L" + 8]).Value2 = enerowin;
            (excelazo.Range["L" + 9]).Value2 = febrerowin;
            (excelazo.Range["L" + 10]).Value2 = marzowin;
            (excelazo.Range["L" + 11]).Value2 = abrilwin;
            (excelazo.Range["L" + 12]).Value2 = mayowin;
            (excelazo.Range["L" + 13]).Value2 = juniowin;
            (excelazo.Range["L" + 14]).Value2 = juliowin;
            (excelazo.Range["L" + 15]).Value2 = agostowin;
            (excelazo.Range["L" + 16]).Value2 = septiembrewin;
            (excelazo.Range["L" + 17]).Value2 = octubrewin;
            (excelazo.Range["L" + 18]).Value2 = noviembrewin;
            (excelazo.Range["L" + 19]).Value2 = diciembrewin;
            (excelazo.Range["M" + 8]).Value2 = eneroloss;
            (excelazo.Range["M" + 9]).Value2 = febreroloss;
            (excelazo.Range["M" + 10]).Value2 = marzoloss;
            (excelazo.Range["M" + 11]).Value2 = abrilloss;
            (excelazo.Range["M" + 12]).Value2 = mayoloss;
            (excelazo.Range["M" + 13]).Value2 = junioloss;
            (excelazo.Range["M" + 14]).Value2 = julioloss;
            (excelazo.Range["M" + 15]).Value2 = agostoloss;
            (excelazo.Range["M" + 16]).Value2 = septiembreloss;
            (excelazo.Range["M" + 17]).Value2 = octubreloss;
            (excelazo.Range["M" + 18]).Value2 = noviembreloss;
            (excelazo.Range["M" + 19]).Value2 = diciembreloss;
            Application.Cells.Locked = false;
            Application.ScreenUpdating = true;
            var rr = 23;
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < listaTrades.Count; i++)
            {
                (excelazo.Range["C" + rr]).Value2 = listaTrades[i].TradeWin;
                (excelazo.Range["D" + rr]).Value2 = listaTrades[i].TradeLoss;
                rr++;
            }
             var rri = 23;
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < listaExito.Count; i++)
            {
                (excelazo.Range["K" + rri]).Value2 = listaExito[i].CantidadExito;
                rri++;
            }
            var rrg = 23;
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < listaFracaso.Count; i++)
            {
                (excelazo.Range["P" + rrg]).Value2 = listaFracaso[i].CantidadFracaso;
                rrg++;
            }

        }
        void Application_SheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            try
            {
                string name = Globals.ThisAddIn.Application.ActiveSheet.Name;
                if (name.Equals(@"ResumenMensual") && target.Column == 1 && target.Row == 4)
                {
                    //var mse = new MensajeDeEspera(() =>
                    //{
                    //    DialogResult continuarCancelacion = MessageBox.Show(@"¿Desea detener la operación?",
                    //    @"Alerta",
                    //    MessageBoxButtons.YesNoCancel,
                    //    MessageBoxIcon.Question);
                    //    cancelar = continuarCancelacion == DialogResult.Yes;
                    //    //return cancelar;
                    //});
                    //mse.Show();
                    var sheet = Globals.ThisAddIn.Application.ActiveSheet;
                    DateTime fecha = DateTime.Now;
                    var fechahoy =fecha.ToOADate();
                    var fechaa = fecha.ToShortDateString();
                    var fechaaa = Convert.ToDateTime(fechaa);
                    var nombredia= (fecha.ToString("dddd", new CultureInfo("es-ES")));
                    var iniciomes = new DateTime(fecha.Year, fecha.Month, 1);//primer dia del mes actual
                    var lunes1 = FirstDayOfWeek(Convert.ToDateTime(iniciomes));
                    var lunes11 = lunes1.ToShortDateString();
                    var martes1 = lunes1.AddDays(1).ToShortDateString();
                    var miercoles1 = lunes1.AddDays(2).ToShortDateString();
                    var jueves1 = lunes1.AddDays(3).ToShortDateString();
                    var viernes1 = lunes1.AddDays(4).ToShortDateString();

                    var lunes2 = lunes1.AddDays(7).ToShortDateString();
                    var martes2 = lunes1.AddDays(8).ToShortDateString();
                    var miercoles2 = lunes1.AddDays(9).ToShortDateString();
                    var jueves2 = lunes1.AddDays(10).ToShortDateString();
                    var viernes2 = lunes1.AddDays(11).ToShortDateString();

                    var lunes3 = lunes1.AddDays(14).ToShortDateString();
                    var martes3 = lunes1.AddDays(15).ToShortDateString();
                    var miercoles3 = lunes1.AddDays(16).ToShortDateString();
                    var jueves3 = lunes1.AddDays(17).ToShortDateString();
                    var viernes3 = lunes1.AddDays(18).ToShortDateString();

                    var lunes4 = lunes1.AddDays(21).ToShortDateString();
                    var martes4 = lunes1.AddDays(22).ToShortDateString();
                    var miercoles4 = lunes1.AddDays(23).ToShortDateString();
                    var jueves4 = lunes1.AddDays(24).ToShortDateString();
                    var viernes4 = lunes1.AddDays(25).ToShortDateString();

                    var lunes5 = lunes1.AddDays(28).ToShortDateString();
                    var martes5 = lunes1.AddDays(29).ToShortDateString();
                    var miercoles5 = lunes1.AddDays(30).ToShortDateString();
                    var jueves5 = lunes1.AddDays(31).ToShortDateString();
                    var viernes5 = lunes1.AddDays(32).ToShortDateString();
                   
                    if (iniciomes == fechaaa)
                    {
                        var nomdia = (iniciomes.ToString("dddd", new CultureInfo("es-ES")));
                        if (nomdia == "lunes")
                        {
                            sheet.Range["B" + 7].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 7].Value2 = fechahoy;
                        }
                        if (nomdia == "martes")
                        {
                            sheet.Range["B" + 8].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 8].Value2 = fechahoy;
                        }
                        if (nomdia == "miércoles")
                        {
                            sheet.Range["B" + 9].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 9].Value2 =fechahoy;     
                        }
                        if (nomdia == "jueves")
                        {
                            sheet.Range["B" + 10].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 10].Value2 = fechahoy;
                        }
                        if (nomdia == "viernes")
                        {
                            sheet.Range["B" + 11].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 11].Value2 = fechahoy;
                        }
                    }
                 
                        if (lunes11 == fechaa)
                        {
                            sheet.Range["B" + 7].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 7].Value2 = fechahoy;

                        }
                        if (lunes2 == fechaa)
                        {
                            sheet.Range["B" + 16].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 16].Value2 = fechahoy;

                        }
                        if (lunes3 == fechaa)
                        {
                            sheet.Range["B" + 25].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 25].Value2 = fechahoy;
                        }
                        if (lunes4 == fechaa)
                        {
                            sheet.Range["B" + 34].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 34].Value2 = fechahoy;
                        }
                        if (lunes5 == fechaa)
                        {
                            sheet.Range["B" + 43].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 43].Value2 = fechahoy;
                        }
                   
                        if (martes1 == fechaa)
                        {
                            sheet.Range["B" + 8].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 8].Value2 = fechahoy;

                        }
                        if (martes2 == fechaa)
                        {
                            sheet.Range["B" + 17].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 17].Value2 = fechahoy;
                        }
                        if (martes3 == fechaa)
                        {
                            sheet.Range["B" + 26].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 26].Value2 = fechahoy;
                        }
                        if (martes4 == fechaa)
                        {
                            sheet.Range["B" + 35].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 35].Value2 = fechahoy;
                        }
                        if (martes5 == fechaa)
                        {
                            sheet.Range["B" + 44].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 44].Value2 = fechahoy;
                        }
                        if (miercoles1 == fechaa)
                        {
                            sheet.Range["B" + 9].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 9].Value2 = fechahoy;

                        }
                        if (miercoles2 == fechaa)
                        {
                            sheet.Range["B" + 18].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 18].Value2 = fechahoy;
                        }
                        if (miercoles3 == fechaa)
                        {
                            sheet.Range["B" + 27].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 27].Value2 = fechahoy;
                        }
                        if (miercoles4 == fechaa)
                        {
                            sheet.Range["B" + 36].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 36].Value2 = fechahoy;
                        }
                        if (miercoles5 == fechaa)
                        {
                            sheet.Range["B" + 45].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 45].Value2 = fechahoy;
                        }
                        if (jueves1 == fechaa)
                        {
                            sheet.Range["B" + 10].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 10].Value2 = fechahoy;

                        }
                        if (jueves2 == fechaa)
                        {
                            sheet.Range["B" + 19].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 19].Value2 = fechahoy;
                        }
                        if (jueves3 == fechaa)
                        {
                            sheet.Range["B" + 28].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 28].Value2 = fechahoy;
                        }
                        if (jueves4 == fechaa)
                        {
                            sheet.Range["B" + 37].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 37].Value2 = fechahoy;
                        }
                        if (jueves5 == fechaa)
                        {
                            sheet.Range["B" + 46].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 46].Value2 = fechahoy;
                        }
                        if (viernes1 == fechaa)
                        {
                            sheet.Range["B" + 11].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 11].Value2 = fechahoy;
                        }
                        if (viernes2 == fechaa)
                        {
                            sheet.Range["B" + 20].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 20].Value2 = fechahoy;
                        }
                        if (viernes3 == fechaa)
                        {
                            sheet.Range["B" + 29].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 29].Value2 = fechahoy;
                        }
                        if (viernes4 == fechaa)
                        {
                            sheet.Range["B" + 38].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 38].Value2 = fechahoy;
                        }
                        if (viernes5 == fechaa)
                        {
                            sheet.Range["B" + 47].NumberFormat = "dd/MM/aaaa";
                            sheet.Range["B" + 47].Value2 = fechahoy;
                        }

                    if (nombredia == "sábado" || nombredia == "domingo")
                    {
                        MessageBox.Show(@"No se puede asignar una fecha, son dias no habiles");
                    }
                    Application.Cells.Locked = false;
                    Application.ScreenUpdating = true;
                }
            }
            catch (Exception e)
            {
                var st = new StackTrace(e, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                MessageBox.Show(e.Message+ line);
            }
        }
}
}
