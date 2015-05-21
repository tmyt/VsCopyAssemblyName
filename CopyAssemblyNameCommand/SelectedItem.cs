namespace tmyt.CopyAssemblyNameCommand
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    
    internal class SelectedItem
    {
        public string Name { get; set; }
        public ItemType Type { get; set; }
    }

    internal enum ItemType
    {
        Assembly,
        Type,
        Namespace
    }
}
