using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTool
{
    public class DataRow
    {
        public DateTime RequestDate { get; set; }
        public string Type { get; set; }
        public string StockCode { get; set; }
        public string StockTarget { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string Region { get; set; }
        public string TruckType { get; set; }
        public int RequestTruckQuantity { get; set; }
        public string FixedTime { get; set; }
        public string PrepareTime { get; set; }
        public string DeliveryTime { get; set; }
        public string Priority { get; set; }
        public string RequestTime { get; set; }
        public string Remark { get; set; }
        public string ProductType { get; set; }
        public string Requester { get; set; }
    }
}
