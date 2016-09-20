namespace System1C.AddIn
{
    /// <summary>
    /// �������� ��� ���������� ������������ ���� ������� ��������������� �������� 1�
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