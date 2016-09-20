using System;
using System.Runtime.InteropServices;

namespace System1C.AddIn
{
    /// <summary>Класс, реализующий пользовательские методы компоненты</summary>
    [ProgId("AddIn.1C_Component")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Component : AddIn, IComponent
    {
        public int Procedure(int parameter)
        {
            throw new NotImplementedException();
        }
    }
}