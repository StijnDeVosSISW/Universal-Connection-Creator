using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCCreator.MODELS
{
    public class BoltDefinition
    {
        // CONSTRUCTOR
        public BoltDefinition()
        {

        }

        // PROPERTIES
        public string Name { get; set; }
        public double ShankDiam { get; set; }
        public double HeadDiam { get; set; }
        public double MaxConnLength { get; set; }
        public string MaterialName { get; set; }
    }
}
