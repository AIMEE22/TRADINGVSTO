using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;

namespace Datos
{
    public class Rest
    {
        public RestClient Cliente;
        public RestRequest Peticion;
        public Rest(string urlApi, string urlMetodo, Method tipo)
        {
            Cliente = new RestClient(urlApi);
            Peticion = new RestRequest(urlMetodo, tipo);
        }
    }
}
