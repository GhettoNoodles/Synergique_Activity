using System.Dynamic;

namespace Synergique_Activity_Formatter.Core
{
    public class Item
    {
        public string Name { get; set; }
        public float[] Sales { get; set; } = new float[12];
        public float AverageSales{ get; set; }
        public string LastInvoiced{ get; set; }
        public string LastPurchased{ get; set; }
        public int CurrentStock;
        public int FormulaDrivenOrder;
        
    }
}