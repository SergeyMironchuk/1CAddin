using System;
using System.Collections;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using EXCEPINFO = System.Runtime.InteropServices.ComTypes.EXCEPINFO;

namespace System1C.AddIn
{
    /// <summary>Функции данного интерфейса вызываются при подключении компоненты</summary>
    /// 
    [Guid("AB634001-F13D-11d0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInitDone
    {
        void Init([MarshalAs(UnmanagedType.IDispatch)] object connection);
        void Done();
        void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info);
    }

    /// <summary>Интерфейс определяет логику вызова функций, процедур и свойств компоненты из 1С</summary>
    /// 
    [Guid("AB634003-F13D-11d0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILanguageExtender
    {
        void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName);
        void GetNProps(ref Int32 props);
        void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum);
        void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName);
        void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
        void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
        void IsPropReadable(Int32 propNum, ref bool propRead);
        void IsPropWritable(Int32 propNum, ref Boolean propWrite);
        void GetNMethods(ref Int32 pMethods);
        void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum);
        void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName);
        void GetNParams(Int32 methodNum, ref Int32 pParams);
        void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue);
        void HasRetVal(Int32 methodNum, ref Boolean retValue);
        void CallAsProc(Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
        void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
    }

    /// <summary>Интерфейс реализован 1С для получения событий от компоненты</summary>
    /// 
    [Guid("AB634004-F13D-11D0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAsyncEvent
    {
        void SetEventBufferDepth(Int32 depth);
        void GetEventBufferDepth(ref long depth);
        void ExternalEvent([MarshalAs(UnmanagedType.BStr)] String source, [MarshalAs(UnmanagedType.BStr)] String message, [MarshalAs(UnmanagedType.BStr)] String data);
        void CleanBuffer();
    }

    /// <summary>С помощью этого интерфейса компонента получает доступ к строке состояния 1С</summary>
    /// 
    [Guid("AB634005-F13D-11D0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStatusLine
    {
        void SetStatusLine([MarshalAs(UnmanagedType.BStr)] String bstrStatusLine);
        void ResetStatusLine();
    }

    [InterfaceType(1)]
    [Guid("3127CA40-446E-11CE-8135-00AA004BB851")]
    public interface IErrorLog
    {
        void AddError(string pszPropName, ref EXCEPINFO pExcepInfo);
    }

    /// <summary>Класс, реализующий свойства и методы для подключения внешней компоненты к 1С</summary>
    public class AddIn : IInitDone, ILanguageExtender
    {
        /// <summary>ProgID COM-объекта компоненты</summary>
        public const string AddInName = "1C_Component";

        // ReSharper disable InconsistentNaming
        public const uint S_OK = 0;
        public const uint S_FALSE = 1;
        public const uint E_POINTER = 0x80004003;
        public const uint E_FAIL = 0x80004005;
        public const uint E_UNEXPECTED = 0x8000FFFF;

        public const short ADDIN_E_NONE = 1000;
        public const short ADDIN_E_ORDINARY = 1001;
        public const short ADDIN_E_ATTENTION = 1002;
        public const short ADDIN_E_IMPORTANT = 1003;
        public const short ADDIN_E_VERY_IMPORTANT = 1004;
        public const short ADDIN_E_INFO = 1005;
        public const short ADDIN_E_FAIL = 1006;
        public const short ADDIN_E_MSGBOX_ATTENTION = 1007;
        public const short ADDIN_E_MSGBOX_INFO = 1008;
        public const short ADDIN_E_MSGBOX_FAIL = 1009;
        // ReSharper restore InconsistentNaming

        /// <summary>Указатель на IDispatch</summary>
        protected object Connect1C;

        /// <summary>Вызов событий 1С</summary>
        protected IAsyncEvent AsyncEvent;

        /// <summary>Статусная строка 1С</summary>
        protected IStatusLine StatusLine;

        protected IErrorLog ErrorLog;

        private Type[] _allInterfaceTypes;  // Коллекция интерфейсов
        private MethodInfo[] _allMethodInfo;  // Коллекция методов
        private PropertyInfo[] _allPropertyInfo; // Коллекция свойств

        private Hashtable _nameToNumberEng; // метод - идентификатор
        private Hashtable _nameToNumberRus; // метод - идентификатор
        private Hashtable _numberToName; // идентификатор - метод
        private Hashtable _numberToParams; // количество параметров метода
        private Hashtable _numberToRetVal; // имеет возвращаемое значение (является функцией)
        private Hashtable _propertyNameToNumberEng; // свойство - идентификатор
        private Hashtable _propertyNameToNumberRus; // свойство - идентификатор
        private Hashtable _propertyNumberToName; // идентификатор - свойство
        private Hashtable _numberToMethodInfoIdx; // номер метода - индекс в массиве методов
        private Hashtable _propertyNumberToPropertyInfoIdx; // номер свойства - индекс в массиве свойств

        /// <summary>
        /// При загрузке 1С:Предприятие V8 инициализирует объект компоненты,
        /// вызывая метод Init и передавая указатель на IDispatch.
        /// Объект может сохранить этот указатель для дальнейшего использования.
        /// Все остальные интерфейсы 1С:Предприятия объект может получить, вызвав метод QueryInterface
        /// переданного ему интерфейса IDispatch. Объект должен возвратить S_OK,
        /// если инициализация прошла успешно, и E_FAIL при возникновении ошибки.
        /// Данный метод может использовать интерфейс IErrorLog для вывода информации об ошибках.
        /// При этом инициализация считается неудачной, если одна из переданных структур EXCEPINFO
        /// имеет поле scode, не равное S_OK. Все переданные в IErrorLog данные обрабатываются
        /// при возврате из данного метода. В момент вызова этого метода свойство AppDispatch не определено.
        /// </summary>
        /// <param name="connection">reference to IDispatch</param>
        public void Init([MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            Connect1C = connection;
            StatusLine = (IStatusLine)connection;
            AsyncEvent = (IAsyncEvent)connection;
            ErrorLog = (IErrorLog)connection;
        }

        /// <summary>
        /// 1С:Предприятие V8 вызывает этот метод при завершении работы с объектом компоненты.
        /// Объект должен возвратить S_OK. Этот метод вызывается независимо от результата
        /// инициализации объекта (метод Init).
        /// </summary>
        public void Done()
        {

        }

        /// <summary>
        /// 1С:Предприятие V8 вызывает этот метод для получения информации о компоненте.
        /// В текущей версии 2.0 компонентной технологии в элемент с индексом 0 необходимо
        /// записать версию поддерживаемой компонентной технологии в формате V_I4 — целого числа,
        /// при этом старший номер версии записывается в тысячные разряды,
        /// младший номер версии — в единицы. Например: версия 3.56 — число 3560.
        /// В настоящее время все объекты внешних компонент могут поддерживать версию 1.0
        /// (соответствует числу 1000) или 2.0 (соответствует 2000).
        /// </summary>
        /// <param name="info">Component information</param>
        public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = 2000;
        }

        /// <summary>Регистрация свойств и методов компоненты в 1C</summary>
        /// <param name="extensionName"></param>
        public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName)
        {
            try
            {
                Type type = GetType();

                _allInterfaceTypes = type.GetInterfaces();
                _allMethodInfo = type.GetMethods();
                _allPropertyInfo = type.GetProperties();

                // Хэш-таблицы с именами методов и свойств компоненты
                _nameToNumberEng = new Hashtable();
                _nameToNumberRus = new Hashtable();
                _numberToName = new Hashtable();
                _numberToParams = new Hashtable();
                _numberToRetVal = new Hashtable();
                _propertyNameToNumberEng = new Hashtable();
                _propertyNameToNumberRus = new Hashtable();
                _propertyNumberToName = new Hashtable();
                _numberToMethodInfoIdx = new Hashtable();
                _propertyNumberToPropertyInfoIdx = new Hashtable();

                int identifier = 0;

                foreach (Type interfaceType in _allInterfaceTypes)
                {
                    // Интересуют только методы в пользовательских интерфейсах, стандартные пропускаем
                    if (interfaceType.Name.Equals("IDisposable")
                        || interfaceType.Name.Equals("IManagedObject")
                        || interfaceType.Name.Equals("IRemoteDispatch")
                        || interfaceType.Name.Equals("IServicedComponentInfo")
                        || interfaceType.Name.Equals("IInitDone")
                        || interfaceType.Name.Equals("ILanguageExtender"))
                    {
                        continue;
                    }

                    // Обработка методов интерфейса
                    MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                    foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                    {
                        var russianName = (RussianNameAttribute)Attribute.GetCustomAttributes(interfaceMethodInfo).FirstOrDefault(a => a is RussianNameAttribute);
                        _nameToNumberEng.Add(interfaceMethodInfo.Name, identifier);
                        _nameToNumberRus.Add(russianName != null ? russianName.Value : interfaceMethodInfo.Name, identifier);
                        _numberToName.Add(identifier, interfaceMethodInfo.Name);
                        _numberToParams.Add(identifier, interfaceMethodInfo.GetParameters().Length);
                        _numberToRetVal.Add(identifier, (interfaceMethodInfo.ReturnType != typeof(void)));
                        identifier++;
                    }

                    // Обработка свойств интерфейса
                    PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                    foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                    {
                        var russianName = (RussianNameAttribute)Attribute.GetCustomAttributes(interfacePropertyInfo).FirstOrDefault(a => a is RussianNameAttribute);
                        _propertyNameToNumberEng.Add(interfacePropertyInfo.Name, identifier);
                        _propertyNameToNumberRus.Add(russianName != null ? russianName.Value : interfacePropertyInfo.Name, identifier);
                        _propertyNumberToName.Add(identifier, interfacePropertyInfo.Name);
                        identifier++;
                    }
                }

                // Отображение номера метода на индекс в массиве
                foreach (DictionaryEntry entry in _numberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < _allMethodInfo.Length; ii++)
                    {
                        if (_allMethodInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            _numberToMethodInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Метод " + entry.Value + " не реализован");
                }

                // Отображение номера свойства на индекс в массиве
                foreach (DictionaryEntry entry in _propertyNumberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < _allPropertyInfo.Length; ii++)
                    {
                        if (_allPropertyInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            _propertyNumberToPropertyInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Свойство " + entry.Value + " не реализовано");
                }

                // Компонент инициализирован успешно. Возвращаем имя компонента.
                extensionName = AddInName;
            }
            catch (Exception)
            {
                // ReSharper disable RedundantJumpStatement
                return;
                // ReSharper restore RedundantJumpStatement
            }
        }

        /// <summary>Возвращает количество свойств</summary>
        /// <param name="props"></param>
        // ReSharper disable once RedundantAssignment
        public void GetNProps(ref Int32 props)
        {
            props = _propertyNameToNumberEng.Count;
        }

        /// <summary>Возвращает целочисленный идентификатор свойства, соответствующий переданному имени</summary>
        /// <param name="propName"></param>
        /// <param name="propNum"></param>
        // ReSharper disable once RedundantAssignment
        public void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum)
        {
            try
            {
                Object result = _propertyNameToNumberEng[propName] ?? _propertyNameToNumberRus[propName];
                propNum = (Int32)result;
            }
            catch (Exception)
            {
                propNum = -1;
            }
        }

        /// <summary>Возвращает имя свойства, соответствующее переданному целочисленному идентификатору</summary>
        /// <param name="propNum"></param>
        /// <param name="propAlias"></param>
        /// <param name="propName"></param>
        // ReSharper disable once RedundantAssignment
        public void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName)
        {
            propName = (String)_propertyNumberToName[propNum];
        }

        /// <summary>Возвращает значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propVal"></param>
        // ReSharper disable once RedundantAssignment
        public void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            propVal = _allPropertyInfo[(int)_propertyNumberToPropertyInfoIdx[propNum]].GetValue(this, null);
        }

        /// <summary>Устанавливает значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propVal"></param>
        public void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            _allPropertyInfo[(int)_propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal, null);
        }

        /// <summary>Определяет, можно ли читать значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propRead"></param>
        // ReSharper disable once RedundantAssignment
        public void IsPropReadable(Int32 propNum, ref bool propRead)
        {
            propRead = _allPropertyInfo[(int)_propertyNumberToPropertyInfoIdx[propNum]].CanRead;
        }

        /// <summary>Определяет, можно ли изменять значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propWrite"></param>
        // ReSharper disable once RedundantAssignment
        public void IsPropWritable(Int32 propNum, ref Boolean propWrite)
        {
            propWrite = _allPropertyInfo[(int)_propertyNumberToPropertyInfoIdx[propNum]].CanWrite;
        }

        /// <summary>Возвращает количество методов</summary>
        /// <param name="pMethods"></param>
        // ReSharper disable once RedundantAssignment
        public void GetNMethods(ref Int32 pMethods)
        {
            pMethods = _nameToNumberEng.Count;
        }

        /// <summary>Возвращает идентификатор метода по его имени</summary>
        /// <param name="methodName">Имя метода</param>
        /// <param name="methodNum">Идентификатор метода</param>
        // ReSharper disable once RedundantAssignment
        public void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum)
        {
            try
            {
                Object result = _nameToNumberEng[methodName] ?? _nameToNumberRus[methodName];
                methodNum = (Int32)result;
            }
            catch (Exception)
            {
                methodNum = -1;
            }
        }

        /// <summary>Возвращает имя метода по идентификатору</summary>
        /// <param name="methodNum"></param>
        /// <param name="methodAlias"></param>
        /// <param name="methodName"></param>
        // ReSharper disable once RedundantAssignment
        public void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName)
        {
            methodName = (String)_numberToName[methodNum];
        }

        /// <summary>Возвращает число параметров метода по его идентификатору</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Число параметров</param>
        // ReSharper disable once RedundantAssignment
        public void GetNParams(Int32 methodNum, ref Int32 pParams)
        {
            pParams = (Int32)_numberToParams[methodNum];
        }

        /// <summary>Получить значение параметра метода по умолчанию</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="paramNum">Номер параметра</param>
        /// <param name="paramDefValue">Возвращаемое значение</param>
        public void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue) { }

        /// <summary>Указывает, что у метода есть возвращаемое значение</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Наличие возвращаемого значения</param>
        // ReSharper disable once RedundantAssignment
        public void HasRetVal(Int32 methodNum, ref Boolean retValue)
        {
            retValue = (Boolean)_numberToRetVal[methodNum];
        }

        /// <summary>Вызов метода как процедуры с использованием идентификатора</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Параметры</param>
        public void CallAsProc(Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                _allMethodInfo[(int)_numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                AsyncEvent.ExternalEvent(AddInName, e.Message, e.ToString());
            }
        }

        /// <summary>Вызов метода как функции с использованием идентификатора</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Возвращаемое значение</param>
        /// <param name="pParams">Параметры</param>
        public void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                retValue = _allMethodInfo[(int)_numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                AsyncEvent.ExternalEvent(AddInName, e.Message, e.ToString());
            }
        }

        /// <summary>
        /// Отправить служебное сообщение 1С
        /// </summary>
        /// <param name="message"></param>
        /// <param name="propName"></param>
        protected void PathMessageToErrorLog(string message, string propName)
        {
            var pExcepInfo = new EXCEPINFO
            {
                bstrDescription = message,
                bstrSource = AddInName,
                pfnDeferredFillIn = (IntPtr)null,
                pvReserved = (IntPtr)null,
                scode = (int)S_OK,
                wCode = ADDIN_E_INFO,
                wReserved = 0,
                bstrHelpFile = "",
                dwHelpContext = 0
            };
            ErrorLog.AddError(propName, ref pExcepInfo);
        }
    }
}