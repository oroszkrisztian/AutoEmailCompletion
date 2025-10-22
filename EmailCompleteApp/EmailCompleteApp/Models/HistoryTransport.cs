using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.Models
{
    public class HistoryTransport
    {
        // Fields public properties
        private int _id;
        private static int _nextId = 1;
        public int NumarComanda { get; set; }
        public string ClientName { get; set; }
        public string CamClient { get; set; }
        public string Route { get; set; }
        public string Transportator { get; set; }
        public DateTime DataTransport { get; set; }

        public HistoryTransport(int numarComanda, string clientName, string camClient, string route, string transportator, DateTime dataTransport)
        {
            _id = _nextId++;
            NumarComanda = numarComanda;
            ClientName = clientName;
            CamClient = camClient;
            Route = route;
            Transportator = transportator;
            DataTransport = dataTransport;
        }

        
    }
}

