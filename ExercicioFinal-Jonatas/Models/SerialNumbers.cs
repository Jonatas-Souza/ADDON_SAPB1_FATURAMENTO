using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExercicioFinal_Jonatas.Models
{
    public class Serialnumber
    {
        public string ManufacturerSerialNumber { get; set; }
        public string InternalSerialNumber { get; set; }
        public string ExpiryDate { get; set; }
        public string ManufactureDate { get; set; }
        public string ReceptionDate { get; set; }
        public string WarrantyStart { get; set; }
        public string WarrantyEnd { get; set; }
        public string Location { get; set; }
        public string Notes { get; set; }
        public string BatchID { get; set; }
        public int SystemSerialNumber { get; set; }
        public int BaseLineNumber { get; set; }
        public decimal? Quantity { get; set; }
        public string TrackingNote { get; set; }
        public string TrackingNoteLine { get; set; }
        public string ItemCode { get; set; }
    }
}
