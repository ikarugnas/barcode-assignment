using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Barcode_Assignment
{
    public class StoreItem
    {
        private double barcode;
        private List<string> fileNames;

        public double Barcode
        {
            get { return barcode; }
            set { barcode = value; }
        }
        public List<string> FileNames
        {
            get { return fileNames; }
            set { fileNames = value; }
        }
    }
}
