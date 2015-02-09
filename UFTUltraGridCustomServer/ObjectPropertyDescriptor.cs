using System;
using System.Collections.Generic;
using System.Text;

namespace UFTUltraGridCustomServer
{
    class ObjectPropertyDescriptor
    {
        public ObjectPropertyDescriptor(string name)
            : this()
        {
            Name = name;
        }

        public ObjectPropertyDescriptor()
        {
            Index = 0;
            IsIndexed = false;
        }

        public string Name { get; set; }
        public bool IsIndexed { get; set; }
        public int Index { get; set; }
    }
}
