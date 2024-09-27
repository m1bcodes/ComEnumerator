using EnumComType;
using System.Runtime.InteropServices;

class Program
{
    [DllImport("ole32.dll")]
    public static extern int CoInitialize(IntPtr pvReserved);

    static void Main(string[] args)
    {
        // Initialize COM
        //int hr = CoInitialize(IntPtr.Zero);
        //if (hr != 0)
        //{
        //    Marshal.ThrowExceptionForHR(hr);
        //}

        HashSet<string> inputTypes = new HashSet<string>();
        HashSet<string> outputTypes = new HashSet<string>();
        HashSet<string> returnTypes = new HashSet<string>();

        TypeLibEnumerator t = new TypeLibEnumerator();

        t.OpenFromID("ATLProject1.ATLSimpleObject.1");

        Console.WriteLine(t.LibDocumentation());

        int nofTypeInfos = t.NofTypeInfos();
        Console.WriteLine("Nof Type Infos: " + nofTypeInfos);

        while (t.NextTypeInfo())
        {
            string typeDoc = t.TypeDocumentation();
            Console.WriteLine();
            Console.WriteLine(typeDoc);
            Console.WriteLine("----------------------------");

            Console.Write("  Interface: ");
            if (t.IsTypeEnum()) Console.Write("Enum");
            if (t.IsTypeRecord()) Console.Write("Record");
            if (t.IsTypeModule()) Console.Write("Module");
            if (t.IsTypeInterface()) Console.Write("Interface");
            if (t.IsTypeDispatch()) Console.Write("Dispatch");
            if (t.IsTypeCoClass()) Console.Write("CoClass");
            if (t.IsTypeAlias()) Console.Write("Alias");
            if (t.IsTypeUnion()) Console.Write("Union");
            if (t.IsTypeMax()) Console.Write("Max");
            Console.WriteLine();

            int nofFunctions = t.NofFunctions();
            int nofVariables = t.NofVariables();
            Console.WriteLine("  functions: " + nofFunctions);
            Console.WriteLine("  variables: " + nofVariables);

            while (t.NextFunction())
            {
                Console.WriteLine();
                Console.WriteLine("  Function     : " + t.FunctionName());

                string funcName = t.FunctionName();
                if (funcName == "Reverse2")
                {
                    int i = 0;
                }

                Console.WriteLine("    returns    : " + t.ReturnType());
                returnTypes.Add(t.ReturnType());

                //Console.Write("    flags      : ");
                //if (t.HasFunctionTypeFlag(TypeLibEnumerator.FDEFAULT)) Console.Write("FDEFAULT ");
                //if (t.HasFunctionTypeFlag(TypeLibEnumerator.FSOURCE)) Console.Write("FSOURCE ");
                //if (t.HasFunctionTypeFlag(TypeLibEnumerator.FRESTRICTED)) Console.Write("FRESTRICTED ");
                //if (t.HasFunctionTypeFlag(TypeLibEnumerator.FDEFAULTVTABLE)) Console.Write("FDEFAULTVTABLE ");
                //Console.WriteLine();

                TypeLibEnumerator.INVOKEKIND ik = t.InvokeKind();
                switch (ik)
                {
                    case TypeLibEnumerator.INVOKEKIND.func:
                        Console.WriteLine("    invoke kind: function");
                        break;
                    case TypeLibEnumerator.INVOKEKIND.put:
                        Console.WriteLine("    invoke kind: put");
                        break;
                    case TypeLibEnumerator.INVOKEKIND.get:
                        Console.WriteLine("    invoke kind: get");
                        break;
                    case TypeLibEnumerator.INVOKEKIND.putref:
                        Console.WriteLine("    invoke kind: put ref");
                        break;
                    default:
                        Console.WriteLine("    invoke kind: ???");
                        break;
                }

                Console.WriteLine("    params     : " + t.NofParameters());
                Console.WriteLine("    params opt : " + t.NofOptionalParameters());

                while (t.NextParameter())
                {
                    Console.Write("    Parameter  : " + t.ParameterName());
                    Console.Write(" type = " + t.ParameterType());
                    if (t.ParameterIsIn())
                    {
                        Console.Write(" in");
                        inputTypes.Add(t.ParameterType());
                    }
                    if (t.ParameterIsOut())
                    {
                        Console.Write(" out");
                        outputTypes.Add(t.ParameterType());
                    }
                    Console.WriteLine();
                }
            }

            while (t.NextVariable())
            {
                Console.WriteLine("  Variable : " + t.VariableName());
                Console.WriteLine("      Type : " + t.VariableType());
                TypeLibEnumerator.VARIABLEKIND vk = t.VariableKind();
                switch (vk)
                {
                    case TypeLibEnumerator.VARIABLEKIND.instance:
                        Console.WriteLine(" (instance)");
                        break;
                    case TypeLibEnumerator.VARIABLEKIND.static_:
                        Console.WriteLine(" (static)");
                        break;
                    case TypeLibEnumerator.VARIABLEKIND.const_:
                        Console.WriteLine(" (const " + t.ConstValue() + ")");
                        break;
                    case TypeLibEnumerator.VARIABLEKIND.dispatch:
                        Console.WriteLine(" (dispatch)");
                        break;
                    default:
                        Console.WriteLine("    variable kind: unknown");
                        break;
                }
            }
        }

        Console.WriteLine("\nUsed input types:");
        foreach (var it in inputTypes) Console.WriteLine(it);

        Console.WriteLine("\nUsed output types:");
        foreach (var it in outputTypes) Console.WriteLine(it);

        Console.WriteLine("\nUsed return types:");
        foreach (var it in returnTypes) Console.WriteLine(it);
    }
}


/*
namespace ComEnumeration
{
    // Define IDispatch interface manually
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00020400-0000-0000-C000-000000000046")]
    public interface IDispatch
    {
        int GetTypeInfoCount(out uint pctinfo);

        int GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo);

        int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames, uint cNames, uint lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

        int Invoke(int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, ref DISPPARAMS pDispParams, out object pVarResult, ref EXCEPINFO pExcepInfo, out uint puArgErr);
    }

    //// Define ITypeInfo interface manually
    //[ComImport]
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //[Guid("00020401-0000-0000-C000-000000000046")]
    //public interface XXITypeInfo
    //{
    //    void GetTypeAttr(out IntPtr ppTypeAttr);
    //    void GetTypeComp(out IntPtr ppTComp);
    //    void GetFuncDesc(uint index, out IntPtr ppFuncDesc);
    //    void GetVarDesc(uint index, out IntPtr ppVarDesc);
    //    void GetNames(int memid, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] StringBuilder[] rgBstrNames, uint cMaxNames, out uint pcNames);
    //    void GetRefTypeOfImplType(uint index, out uint pRefType);
    //    void GetImplTypeFlags(uint index, out int pImplTypeFlags);
    //    void GetIDsOfNames(string[] rgszNames, uint cNames, int[] pMemId);
    //    void Invoke();
    //    void GetDocumentation(int memid, out string pBstrName, out string pBstrDocString, out uint pdwHelpContext, out string pBstrHelpFile);
    //    void GetDllEntry();
    //    void GetRefTypeInfo();
    //    void AddressOfMember();
    //    void CreateInstance();
    //    void GetMops();
    //    void GetContainingTypeLib();
    //    void ReleaseTypeAttr(IntPtr pTypeAttr);
    //    void ReleaseFuncDesc(IntPtr pFuncDesc);
    //    void ReleaseVarDesc(IntPtr pVarDesc);
    //}

    // Define DISPPARAMS structure
    //[StructLayout(LayoutKind.Sequential)]
    //public struct DISPPARAMS
    //{
    //    public IntPtr rgvarg;
    //    public IntPtr rgdispidNamedArgs;
    //    public uint cArgs;
    //    public uint cNamedArgs;
    //}

    //// Define EXCEPINFO structure
    //[StructLayout(LayoutKind.Sequential)]
    //public struct EXCEPINFO
    //{
    //    public ushort wCode;
    //    public ushort wReserved;
    //    [MarshalAs(UnmanagedType.BStr)] public string bstrSource;
    //    [MarshalAs(UnmanagedType.BStr)] public string bstrDescription;
    //    [MarshalAs(UnmanagedType.BStr)] public string bstrHelpFile;
    //    public uint dwHelpContext;
    //    public IntPtr pvReserved;
    //    public IntPtr pfnDeferredFillIn;
    //    public int scode;
    //}

    public class ComReflection
    {
        public static void EnumerateComMethods(IDispatch dispatch)
        {
            // Get type information count
            dispatch.GetTypeInfoCount(out uint typeInfoCount);
            if (typeInfoCount == 0)
            {
                Console.WriteLine("No type information available.");
                return;
            }

            // Get type information
            dispatch.GetTypeInfo(0, 0, out ITypeInfo typeInfo);

            // Get type attributes
            typeInfo.GetTypeAttr(out IntPtr typeAttrPtr);
            TYPEATTR typeAttr = Marshal.PtrToStructure<TYPEATTR>(typeAttrPtr);

            typeInfo.GetContainingTypeLib(out ITypeLib ppTLB, out int pIndex);
            int numOfTypeInfos = ppTLB.GetTypeInfoCount();
            for (int iti = 0; iti < numOfTypeInfos; iti++)
            {
                ppTLB.GetTypeInfo(iti, out ITypeInfo ppTI);
                ppTI.GetTypeAttr(out IntPtr ppTypeAttr);
                typeAttr = Marshal.PtrToStructure<TYPEATTR>(typeAttrPtr);
                ppTI.GetDocumentation(-1, out string Name, out string DocString, out int dwHelpCtx, out string helpFile);
                Console.WriteLine(Name);

            }
            Console.WriteLine($"Number of functions: {typeAttr.cFuncs}");

            // Enumerate functions
            for (int i = 0; i < typeAttr.cFuncs; i++)
            {
                //typeInfo.GetDocumentation((int)i, out string name, out string doc, out int hc, out string helpFile);

                typeInfo.GetFuncDesc(i, out IntPtr funcDescPtr);
                FUNCDESC funcDesc = Marshal.PtrToStructure<FUNCDESC>(funcDescPtr);

                // Get function name
                string[] names = new string[256];
                typeInfo.GetNames(funcDesc.memid, names, 256, out int nameCount);
                Console.WriteLine($"Function {i}:");
                for (int j = 0; j < nameCount; j++)
                    Console.WriteLine("   " + names[j]);

                typeInfo.ReleaseFuncDesc(funcDescPtr);
            }

            typeInfo.ReleaseTypeAttr(typeAttrPtr);
        }
    }

    //[StructLayout(LayoutKind.Sequential)]
    //public struct TYPEATTR
    //{
    //    public Guid guid;
    //    public uint lcid;
    //    public uint dwReserved;
    //    public int memidConstructor;
    //    public int memidDestructor;
    //    public IntPtr lpstrSchema;
    //    public uint cbSizeInstance;
    //    public TYPEKIND typekind;
    //    public ushort cFuncs;
    //    public ushort cVars;
    //    public ushort cImplTypes;
    //    public ushort cbSizeVft;
    //    public ushort cbAlignment;
    //    public ushort wTypeFlags;
    //    public ushort wMajorVerNum;
    //    public ushort wMinorVerNum;
    //    public TYPEDESC tdescAlias;
    //    public IDLDESC idldescType;
    //}

    //[StructLayout(LayoutKind.Sequential)]
    //public struct FUNCDESC
    //{
    //    public int memid;
    //    public IntPtr lprgscode;
    //    public IntPtr lprgelemdescParam;
    //    public FUNCKIND funckind;
    //    public INVOKEKIND invkind;
    //    public CALLCONV callconv;
    //    public short cParams;
    //    public short cParamsOpt;
    //    public short oVft;
    //    public short cScodes;
    //    public ELEMDESC elemdescFunc;
    //    public ushort wFuncFlags;
    //}

    class Program
    {
        public static void Main(string[] args)
        {
            string progId = "ATLProject1.ATLSimpleObject.1";
            Object so = Activator.CreateInstance(Type.GetTypeFromProgID(progId));
            ComReflection.EnumerateComMethods((IDispatch)so);
        }
    }
}
*/