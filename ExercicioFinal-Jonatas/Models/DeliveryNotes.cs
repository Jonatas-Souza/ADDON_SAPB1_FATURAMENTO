﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExercicioFinal_Jonatas.Models
{
    public class DeliveryNotes : Documents
    {
        public DeliveryNotes()
        {
            DocObjectCode = "oDeliveryNotes";
            DocumentLines = new List<DocumentLine>();
            DocumentReferences = new List<DocumentReference>();
        }
        public List<DocumentLine> DocumentLines { get; set; }
        public TaxExtension TaxExtension { get; set; }
        public AddressExtension AddressExtension { get; set; }
        public List<DocumentReference> DocumentReferences { get; set; }
    }
}
