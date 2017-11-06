using System;
using System.Collections.Generic;


namespace Respuesta
{
    public class General
    {
        public DateTime FechaIni { get; set; }
        public DateTime FechaFin { get; set; }
        public General()
        {
        }
    }
    public class DatosSemanal
    {
        public string Id { get; set; }
        public string Dia { get; set;  }
        public  string Fecha { get; set; }
        public  string TipoTrade { get; set; }
        public  string Proceso { get; set; }
        public string Direccion { get; set; }
        public  double  Puntos { get; set; }
        public  string Exito { get; set; }
        public  string Fracaso { get; set; }
        public List<DatosSemanal> ListaReSemanals { get; set; }
    }

    public class DatosAnual
    {
        public string Mes { get; set; }
        public string Fecha { get; set; }
        public string Trades { get; set; }
        public int Win { get; set; }
        public int Loss { get; set; }
        public int PuntosWin { get; set; }
        public int PuntosLoss { get; set; }
        public double PuntosTotales { get; set; }
        public string TradeWin { get; set; }
        public string TradeLoss { get; set; }
        public int Total { get; set; }
        public int Largo { get; set; }
        public int Corto { get; set; }
        public int CantidadExito { get; set; }
        public int CantidadFracaso { get; set; }
        public List<DatosAnual> ListaAnualList { get; set; }
    }

    public class Guardar
    {
        public string Fecha { get; set; }
        public string TipoTrade { get; set; }
        public string Proceso { get; set; }
        public string Direccion { get; set; }
        public double Puntos { get; set; }
        public string Exito { get; set; }
        public string Fracaso { get; set; }
        public string Id { get; set; }
    }
}
