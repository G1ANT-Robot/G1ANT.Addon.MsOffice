namespace G1ANT.Addon.MSOffice.Models.Access
{
    internal class NameValueModel : INameModel
    {
        public string Name { get; set; }
        public object Value { get; set; }

        public NameValueModel(string name, object value)
        {
            Name = name;
            Value = value;
        }
    }
}
