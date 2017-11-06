using System;
using System.Collections.Generic;
using System.Net;
using Herramienta;
using Respuesta;
using RestSharp;


namespace Datos
{
   public class Reporte
    {
        public static void AvanceSemanal(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Semanal.ResumenSemana, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }

        public static void ReporteAnual(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Anual.ResumenAnual, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }

        public static void SeleccionarTradeGanador(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Anual.TradeWin, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }
        public static void SeleccionarTradePerdedor(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Anual.TradeLoss, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }
        public static void TradesAgrupados(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Anual.TradesAgrupados, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }
        public static void SeleccionarRazonExito(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Anual.RazonExitoGroup, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }
        public static void SeleccionarRazonFracaso(Action<IRestResponse> callback, General fechas)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Anual.RazonFracasoGroup, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(fechas);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            callback(null);
                            break;
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResSemanal, "EXCEPCION: " + e.Message);
            }
        }
        public static void Guardar(Action<IRestResponse> callback, List<Guardar> guardar)
        {
            try
            {
                var rest = new Rest(Local.Api.UrlApi, Resumen.Semanal.GuardarResumen, Method.POST);
                rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
                rest.Peticion.AddJsonBody(guardar);
                rest.Cliente.ExecuteAsync(rest.Peticion, response =>
                {
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            callback(response);
                            break;
                        default:
                            throw new Exception(@"error al buscar articulo");
                    }
                });
            }
            catch (Exception e)
            {
                Opcion.Log(Log.Interno.ResMensual, "EXCEPCION: " + e.Message);
                // callback("CONTINUAR");
            }
        }
        public static void Guardado(Action<IRestResponse> callback, Guardar lista)
        {
            var rest = new Rest(Local.Api.UrlApi, Resumen.Semanal.GuardarResumen, Method.POST);
            rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
            rest.Peticion.AddJsonBody(lista);
            rest.Cliente.ExecuteAsync(rest.Peticion, response =>
            {
                switch (response.StatusCode)
                {
                    case HttpStatusCode.OK:
                        callback(response);
                        break;
                    default:
                        throw new Exception(@"error al buscar articulo");
                }
            });
        }
        public static void Actualizar(Guardar actualiza)
        {
            var rest = new Rest(Local.Api.UrlApi, Resumen.Semanal.ActualizarResumen, Method.POST);
            rest.Peticion.AddHeader(Constantes.Http.ObtenerTipoDeContenido, Constantes.Http.TipoDeContenido.Json);
            rest.Peticion.AddJsonBody(actualiza);
            rest.Cliente.ExecuteAsync(rest.Peticion, response =>
            {
                switch (response.StatusCode)
                {
                    case HttpStatusCode.OK:
                        break;
                    default:
                        throw new Exception(@"Los datos no se pudieron actualizar");
                }
            });
        }

    }
}

