using System.Runtime.InteropServices;

namespace System1C.AddIn
{
    /// <summary>»нтерфейс с объ€влени€ми пользовательских методов и свойств компоненты</summary>
    /// 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IComponent
    {
        [RussianName("ћетодЌа–усскомязыке")]
        int Procedure(int parameter);
    }
}