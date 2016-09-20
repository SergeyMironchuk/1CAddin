namespace System1C.AddIn
{
    /// <summary>
    /// Аттрибут для присвоения рускоязычных имен методам разрабатываемыз плагинов 1С
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.All)]
    public class RussianNameAttribute : System.Attribute
    {
        private readonly string _russianName;
        public RussianNameAttribute(string russianName)
        {
            _russianName = russianName;
        }

        public string Value
        {
            get { return _russianName; }
        }

        public override string ToString()
        {
            return _russianName;
        }
    }
}