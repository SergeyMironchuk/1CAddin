using System.Runtime.InteropServices;

namespace System1C.AddIn
{
    /// <summary>��������� � ������������ ���������������� ������� � ������� ����������</summary>
    /// 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IComponent
    {
        [RussianName("�������������������")]
        int Procedure(int parameter);
    }
}