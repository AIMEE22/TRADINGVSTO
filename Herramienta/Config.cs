using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Herramienta
{
  public static class Local
    {
        public static class Api
        {
            public static string UrlLocal { get; set; } = "http://192.168.0.36:8080/apimovstock/phptrading/public/";
            public static string UrlApi { get; set; } = Properties.Settings.Default.UrlApi;
        }
    }
    public static class Log
    {
        public static class Interno
        {
            public static string ResSemanal { get; set; } = "semana.log";
            public static string ResMensual { get; set; } = "mes.log";
            public static string ResAnual { get; set; } = "año.log";
        }


    }
    public static class Resumen
    {
        public static class Semanal
        {
            public static string ResumenSemana { get; set; } = Properties.Settings.Default.ResumenSemanal;
            public static string GuardarResumen { get; set; } = Properties.Settings.Default.Insertar;
            public static string ActualizarResumen { get; set; } = Properties.Settings.Default.Actualizar;

        }

        public static class Anual
        {
            public static string ResumenAnual { get; set; } = Properties.Settings.Default.ReseumenAnual;
            public static string TradeWin { get; set; } = Properties.Settings.Default.TradeGanador;
            public static string TradeLoss { get; set; } = Properties.Settings.Default.TradePerdedor;
            public static string TradesAgrupados { get; set; } = Properties.Settings.Default.TradesAgrupados;
            public static string RazonExitoGroup { get; set; } = Properties.Settings.Default.RazonExito;
            public static string RazonFracasoGroup { get; set; } = Properties.Settings.Default.RazonFracaso;

        }
    }
  
}
 
           