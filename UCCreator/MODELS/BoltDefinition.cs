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
        public int ShankDiam { get; set; }
        public int HeadDiam { get; set; }
        public int MaxConnLength { get; set; }
        public string MaterialName { get; set; }
    }
}
