using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Datos;
using Herramienta;
using Respuesta;
using Office = Microsoft.Office.Core;

// TODO:  Siga estos pasos para habilitar el elemento (XML) de la cinta de opciones:

// 1: Copie el siguiente bloque de código en la clase ThisAddin, ThisWorkbook o ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Cree métodos de devolución de llamada en el área "Devolución de llamadas de la cinta de opciones" de esta clase para controlar acciones del usuario,
//    como hacer clic en un botón. Nota: si ha exportado esta cinta de opciones desde el diseñador de la cinta de opciones,
//    mueva el código de los controladores de eventos a los métodos de devolución de llamada y modifique el código para que funcione con el
//    modelo de programación de extensibilidad de la cinta de opciones (RibbonX).

// 3. Asigne atributos a las etiquetas de control del archivo XML de la cinta de opciones para identificar los métodos de devolución de llamada apropiados en el código.  

// Para obtener más información, vea la documentación XML de la cinta de opciones en la Ayuda de Visual Studio Tools para Office.


namespace TRADINGVSTO
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        public Ribbon1()
        {
        }

        #region Miembros de IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("TRADINGVSTO.Ribbon1.xml");
        }

        #endregion

        #region Devoluciones de llamada de la cinta de opciones
        //Cree aquí métodos de devolución de llamada. Para obtener más información sobre los métodos de devolución de llamada, visite http://go.microsoft.com/fwlink/?LinkID=271226.

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        #endregion

        #region Aplicaciones auxiliares

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        return resourceReader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion
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

        public List<DatosSemanal> ListaDatosSemanals;
        public void ResumenMensual(Office.IRibbonControl control)
        {
            var addIn = Globals.ThisAddIn;
            DateTime fecha = DateTime.Now;
           var iniciomes = new DateTime(fecha.Year, fecha.Month, 1).ToShortDateString();//primer dia del mes actual
           var finmes = new DateTime(fecha.Year, fecha.Month+1, 1).AddDays(-1).ToShortDateString();//ultimo dia del mes actual
            //var inicio = Convert.ToDateTime(iniciomes).ToString("yyyy-MM-dd");
            //var fin = Convert.ToDateTime(finmes).ToString("yyyy-MM-dd");
            Opcion.EjecucionAsync(x =>
            {
                var times = new General
                {
                    FechaFin = Convert.ToDateTime(finmes),
                    FechaIni = Convert.ToDateTime(iniciomes)
                };
                Reporte.AvanceSemanal(x, times);
            }, jsonResult =>
            {
                if (jsonResult != null)
                {
                 var listaSemana = Opcion.JsonaListaGenerica<DatosSemanal>(jsonResult).ToList();
                    ListaDatosSemanals = listaSemana;
                 addIn.ResumenSemanal(listaSemana);
                }
                else
                {
                    MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                }
            });
        }
        public void ReporteAnual(Office.IRibbonControl control)
        {
            var addIn = Globals.ThisAddIn;
            
            var fecha = DateTime.Now.ToShortDateString();
            var value = fecha;
            const char delimite = '/';
            var substrings = value.Split(delimite);
            var year = substrings[2];
            var fechainicial ="01/01/"+ year;
            var fechafinal ="12/31/"+ year;
            var date1 = DateTime.Parse(fechainicial,CultureInfo.InvariantCulture);
            var date2 = DateTime.Parse(fechafinal,CultureInfo.InvariantCulture);
            var datosAnual=new List<DatosAnual>();
            var datosGanador=new List<DatosAnual>();
            var datosPerdedor=new List<DatosAnual>();
            var datosTrade= new List<DatosAnual>();
            var datosExito= new List<DatosAnual>();
            List<DatosAnual> datosFracaso;
            var times = new General
            {
                FechaIni = Convert.ToDateTime(date1),
                FechaFin = Convert.ToDateTime(date2)
            };
            Opcion.EjecucionAsync(x =>
            {
                Reporte.ReporteAnual(x, times);
            }, jsonResult =>
            {
                if (jsonResult != null)
                {
                    datosAnual = Opcion.JsonaListaGenerica<DatosAnual>(jsonResult).ToList();
                }
                else
                {
                    MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                }
          
            Opcion.EjecucionAsync(y =>
            {
                Reporte.SeleccionarTradeGanador(y, times);
            }, jsonResu =>
            {
                if (jsonResu != null)
                {
                    datosGanador = Opcion.JsonaListaGenerica<DatosAnual>(jsonResu).ToList();
                }
                else
                {
                    MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                }
            Opcion.EjecucionAsync(z =>
            {
                Reporte.SeleccionarTradePerdedor(z, times);
            }, jsonR =>
            {
                if (jsonR != null)
                {
                    datosPerdedor = Opcion.JsonaListaGenerica<DatosAnual>(jsonR).ToList();
                }
                else
                {
                    MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                }
                Opcion.EjecucionAsync(z =>
                {
                    Reporte.TradesAgrupados(z, times);
                }, jsonRe =>
                {
                    if (jsonRe != null)
                    {
                        datosTrade = Opcion.JsonaListaGenerica<DatosAnual>(jsonRe).ToList();
                    }
                    else
                    {
                        MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                    }
                    Opcion.EjecucionAsync(z =>
                    {
                        Reporte.SeleccionarRazonExito(z, times);
                    }, json =>
                    {
                        if (json != null)
                        {
                            datosExito = Opcion.JsonaListaGenerica<DatosAnual>(json).ToList();
                        }
                        else
                        {
                            MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                        }
                        Opcion.EjecucionAsync(z =>
                        {
                            Reporte.SeleccionarRazonFracaso(z, times);
                        }, jso =>
                        {
                            if (jso != null)
                            {
                                datosFracaso = Opcion.JsonaListaGenerica<DatosAnual>(jso).ToList();
                                addIn.ReporteAnual(datosAnual, datosGanador, datosPerdedor,datosTrade,datosExito,datosFracaso);
                            }
                            else
                            {
                                MessageBox.Show(@"No se encontro informacion con los paramentro de busqueda");
                            }
                        });
                    });
                });
            });
         });         
        });
        }

        public string Dato;
        public string Dato2;
        // ReSharper disable once FunctionComplexityOverflow
        public void GuardarDatos(Office.IRibbonControl control)
        {
            try
            {
                string name = Globals.ThisAddIn.Application.ActiveSheet.Name;
                if (name.Equals(@"ResumenMensual"))
                {
                    //var mse = new MensajeDeEspera(() =>
                    //{
                    //    DialogResult continuarCancelacion = MessageBox.Show(@"¿Desea detener la operación?",
                    //    @"Alerta",
                    //    MessageBoxButtons.YesNoCancel,
                    //    MessageBoxIcon.Question);
                    //    cancelar = continuarCancelacion == DialogResult.Yes;
                    //    return cancelar;
                    //});
                    //mse.Show();
                    var sheet = Globals.ThisAddIn.Application.ActiveSheet;
                    //var nInLastRow = sheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByRows,
                    //   XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value)
                    //    .Row;
                    object[,] value = sheet.Range["B7", "M11"].Value;
                    object[,] value1 = sheet.Range["B16", "M20"].Value;
                    object[,] value2= sheet.Range["B25", "M29"].Value;
                    object[,] value3= sheet.Range["B34", "M38"].Value;
                    object[,] value4 = sheet.Range["B43", "M47"].Value;
                    var lista = new List<Guardar>();
                    for (var x = 1; x <= (value.Length / 12); x++)
                    {
                        Dato = value[x, 10]== null ? "" : value[x, 10].ToString();
                        Dato2 = value[x, 11]==null ? "" : value[x, 11].ToString();
                        var id = value[x, 12] == null ? "" : value[x, 12].ToString(); 
                        var fechaguardar = Convert.ToDateTime(value[x, 1]).ToString("yyyy-MM-dd");
                        {
                            if (value[x, 1] != null && value[x, 2] != null && value[x, 3] != null && value[x, 5] != null &&
                                value[x, 6] != null)
                            {
                                lista.Add(new Guardar
                                {
                                    Fecha = fechaguardar,
                                    TipoTrade = value[x, 2].ToString(),
                                    Proceso = value[x, 3].ToString(),
                                    Direccion = value[x, 5].ToString(),
                                    Puntos = Convert.ToDouble(value[x, 6]),
                                    Exito = Dato,
                                    Fracaso = Dato2,
                                    Id =id
                                });
                            }   
                        }   
                    }
                    for (var x = 1; x <= (value1.Length / 12); x++)
                    {
                        Dato = value1[x, 10] == null ? "" : value1[x, 10].ToString();
                        Dato2 = value1[x, 11] == null ? "" : value1[x, 11].ToString();
                        var id = value1[x, 12] == null ? "" : value1[x, 12].ToString();
                        var fechaguardar = Convert.ToDateTime(value1[x, 1]).ToString("yyyy-MM-dd");
                        {
                            if (value1[x, 1] != null && value1[x, 2] != null && value1[x, 3] != null && value1[x, 5] != null &&
                                value1[x, 6] != null)
                            {
                                lista.Add(new Guardar
                                {
                                    Fecha = fechaguardar,
                                    TipoTrade = value1[x, 2].ToString(),
                                    Proceso = value1[x, 3].ToString(),
                                    Direccion = value1[x, 5].ToString(),
                                    Puntos = Convert.ToDouble(value1[x, 6]),
                                    Exito = Dato,
                                    Fracaso = Dato2,
                                    Id = (id)
                                });
                            }
                        }
                    }
                    for (var x = 1; x <= (value2.Length / 12); x++)
                    {
                        Dato = value2[x, 10] == null ? "" : value2[x, 10].ToString();
                        Dato2 = value2[x, 11] == null ? "" : value2[x, 11].ToString();
                        var id = value2[x, 12] == null ? "" : value2[x, 12].ToString();
                        var fechaguardar = Convert.ToDateTime(value2[x, 1]).ToString("yyyy-MM-dd");
                        {

                            if (value2[x, 1] != null && value2[x, 2] != null && value2[x, 3] != null && value2[x, 5] != null &&
                                value2[x, 6] != null)
                            {
                                lista.Add(new Guardar
                                {
                                    Fecha = fechaguardar,
                                    TipoTrade = value2[x, 2].ToString(),
                                    Proceso = value2[x, 3].ToString(),
                                    Direccion = value2[x, 5].ToString(),
                                    Puntos = Convert.ToDouble(value2[x, 6]),
                                    Exito = Dato,
                                    Fracaso = Dato2,
                                    Id = (id)
                                });
                            }
                        }
                    }
                    for (var x = 1; x <= (value3.Length / 12); x++)
                    {
                        Dato = value3[x, 10] == null ? "" : value3[x, 10].ToString();
                        Dato2 = value3[x, 11] == null ? "" : value3[x, 11].ToString();
                        var id = value3[x, 12] == null ? "" : value3[x, 12].ToString();
                        var fechaguardar = Convert.ToDateTime(value3[x, 1]).ToString("yyyy-MM-dd");
                        {
                            if (value3[x, 1] != null && value3[x, 2] != null && value3[x, 3] != null && value3[x, 5] != null &&
                                value3[x, 6] != null)
                            {
                                lista.Add(new Guardar
                                {
                                    Fecha = fechaguardar,
                                    TipoTrade = value3[x, 2].ToString(),
                                    Proceso = value3[x, 3].ToString(),
                                    Direccion = value3[x, 5].ToString(),
                                    Puntos = Convert.ToDouble(value3[x, 6]),
                                    Exito = Dato,
                                    Fracaso = Dato2,
                                    Id = id
                                });
                            }
                        }
                    }
                    for (var x = 1; x <= (value4.Length / 12); x++)
                    {
                        Dato = value4[x, 10] == null ? "" : value4[x, 10].ToString();
                        Dato2 = value4[x, 11] == null ? "" : value4[x, 11].ToString();
                        var fechaguardar = Convert.ToDateTime(value4[x, 1]).ToString("yyyy-MM-dd");
                        var id = value4[x, 12] == null ? "" : value4[x, 12].ToString();
                        {
                            if (value4[x, 1] != null && value4[x, 2] != null && value4[x, 3] != null && value4[x, 5] != null &&
                                value4[x, 6] != null)
                            {
                                lista.Add(new Guardar
                                {
                                    Fecha = fechaguardar,
                                    TipoTrade = value4[x, 2].ToString(),
                                    Proceso = value4[x, 3].ToString(),
                                    Direccion = value4[x, 5].ToString(),
                                    Puntos = Convert.ToDouble(value4[x, 6]),
                                    Exito = Dato,
                                    Fracaso = Dato2,
                                    Id =id
                                });
                            }
                        }
                    }
                    // ReSharper disable once ForCanBeConvertedToForeach
                    for (var i = 0; i < lista.Count; i++)
                    {
                            if (lista[i].Id=="" && lista[i].Fecha!=null && lista[i].TipoTrade != null && lista[i].Proceso != null &&
                                lista[i].Direccion != null)
                        {
                            var listaguardar = new Guardar
                            {
                                Fecha = lista[i].Fecha,
                                TipoTrade = lista[i].TipoTrade,
                                Proceso = lista[i].Proceso,
                                Direccion = lista[i].Direccion,
                                Puntos = lista[i].Puntos,
                                Exito = lista[i].Exito,
                                Fracaso = lista[i].Fracaso
                            };
                            Opcion.EjecucionAsync(x =>
                            {

                                Reporte.Guardado(x, listaguardar);
                            }, resultado =>
                            {
                            });
                             }
                            else if((lista[i].Id !=null && lista[i].Fecha != null && lista[i].TipoTrade != null && lista[i].Proceso != null &&
                                     lista[i].Direccion != null))
                            {
                            var listaguardar = new Guardar
                            {
                                Id=lista[i].Id,
                                Fecha = lista[i].Fecha,
                                TipoTrade = lista[i].TipoTrade,
                                Proceso = lista[i].Proceso,
                                Direccion = lista[i].Direccion,
                                Puntos = lista[i].Puntos,
                                Exito = lista[i].Exito,
                                Fracaso = lista[i].Fracaso
                            };
                             Reporte.Actualizar(listaguardar);
                            }
                    }
                    MessageBox.Show(@"La informacion se a guardado correctamente");
                }
                else
                {
                    throw new Exception(
                        @"Debes escoger la hoja de trabajo 'ResumenMensual' para seleccionar esta opción.");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
    }
