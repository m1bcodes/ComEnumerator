//MIT License

//Copyright (c) 2024 m1bcodes

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.


using ComTypeHelper;
using System.Runtime.InteropServices;

class Program
{
    [DllImport("ole32.dll")]
    public static extern int CoInitialize(IntPtr pvReserved);

    static void Main(string[] args)
    {
        string progId = "ATLProject1.ATLSimpleObject.1";
        // Initialize COM
        //int hr = CoInitialize(IntPtr.Zero);
        //if (hr != 0)
        //{
        //    Marshal.ThrowExceptionForHR(hr);
        //}

        HashSet<string> inputTypes = new HashSet<string>();
        HashSet<string> outputTypes = new HashSet<string>();
        HashSet<string> returnTypes = new HashSet<string>();

        TypeLibEnumerator t = new TypeLibEnumerator(progId);

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

        t = new TypeLibEnumerator(progId);
        var ed = t.EnumDictionary();

        dynamic enumContainer = ComContainer.createEnumContainer(ed);
        Console.WriteLine("FocusPolicy.StrongFocus = {0}", enumContainer.FocusPolicy.StrongFocus);

        dynamic api = ComContainer.CreateComWrapper("api", progId);
        Console.WriteLine("api.FocusPolicy.StrongFocus = {0}", api.FocusPolicy.StrongFocus);
        Console.WriteLine("api.Reverse2 = {0}", api.api.Reverse2("Hallo"));
    }
}
