namespace CopyAssemblyNameCommand
{
    
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
