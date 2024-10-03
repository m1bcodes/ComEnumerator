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

using System.Reflection.Emit;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComTypeHelper
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

    public class TypeLibHelper
    {
        public static string VartypeToString(VarEnum vartype)
        {
            switch (vartype)
            {
                case VarEnum.VT_EMPTY:
                    return "Empty";
                case VarEnum.VT_NULL:
                    return "Null";
                case VarEnum.VT_I2:
                    return "Short";
                case VarEnum.VT_I4:
                    return "Long";
                case VarEnum.VT_R4:
                    return "Float";
                case VarEnum.VT_R8:
                    return "Double";
                case VarEnum.VT_CY:
                    return "Currency";
                case VarEnum.VT_DATE:
                    return "Date";
                case VarEnum.VT_BSTR:
                    return "BSTR";
                case VarEnum.VT_BOOL:
                    return "Boolean";
                case VarEnum.VT_DECIMAL:
                    return "Decimal";
                case VarEnum.VT_I1:
                    return "Char";
                case VarEnum.VT_UI1:
                    return "Unsigned Char";
                case VarEnum.VT_UI2:
                    return "Unsigned Short";
                case VarEnum.VT_UI4:
                    return "Unsigned Long";
                case VarEnum.VT_I8:
                    return "int64";
                case VarEnum.VT_UI8:
                    return "uint64";
                case VarEnum.VT_INT:
                    return "Int";
                case VarEnum.VT_UINT:
                    return "Unsigned Int";
                case VarEnum.VT_VOID:
                    return "Void";
                case VarEnum.VT_HRESULT:
                    return "HRESULT";
                case VarEnum.VT_UNKNOWN:
                    return "Unknown";
                case VarEnum.VT_DISPATCH:
                    return "Dispatch";
                case VarEnum.VT_ARRAY:
                    return "Array";
                case VarEnum.VT_BYREF:
                    return "ByRef";
                default:
                    return "Unknown VARTYPE (" + vartype + ")";
            }
        }
    }

    public abstract class TypeInfo
    {
        new abstract public string ToString();
    }

    public class ScalarTypeInfo : TypeInfo
    {
        private VarEnum vt;

        public ScalarTypeInfo(VarEnum vt)
        {
            this.vt = vt;
        }

        public override string ToString()
        {
            return TypeLibHelper.VartypeToString(vt);
        }
    }

    public class UserTypeInfo : TypeInfo
    {
        private string typeName;

        public UserTypeInfo(string name)
        {
            this.typeName = name;
        }

        public override string ToString()
        {
            return $"User type ({typeName})";
        }
    }

    public class PointerTypeInfo : TypeInfo
    {
        private TypeInfo child;

        public PointerTypeInfo(TypeInfo child)
        {
            this.child = child;
        }

        public override string ToString()
        {
            return $"Pointer({child.ToString()})";
        }
    }

    public class SafeArrayTypeInfo : TypeInfo
    {
        private TypeInfo elemType;

        public SafeArrayTypeInfo(TypeInfo elemType)
        {
            this.elemType = elemType;
        }

        public override string ToString()
        {
            return $"SafeArray({elemType.ToString()})";
        }
    }

    public class UnknownTypeInfo : TypeInfo
    {
        private string description;

        public UnknownTypeInfo(string desc)
        {
            this.description = desc;
        }

        public override string ToString()
        {
            return $"Unknown type ({description})";
        }
    }

    public class TypeLibEnumerator
    {
        public TypeLibEnumerator()
        {
           // OleInitialize(IntPtr.Zero);
        }

        [DllImport("Ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern void OleInitialize(IntPtr pvReserved);

        public void OpenFromID(string progId)
        {
            Object so = Activator.CreateInstance(Type.GetTypeFromProgID(progId));
            IDispatch dispatch = (IDispatch)so;

            // Get type information
            dispatch.GetTypeInfo(0, 0, out ITypeInfo typeInfo);

            typeInfo.GetContainingTypeLib(out this.ppTLB, out int pIndex);
            this.typeInfoCount = ppTLB.GetTypeInfoCount();
        }

        public bool NextTypeInfo()
        {
            this.curTypeInfo++;

            this.curFunc = -1;
            this.curVar = -1;

            if (this.curTypeInfo >= this.typeInfoCount) return false;

            if (this.curTypeAttrPtr != IntPtr.Zero && this.curITypeInfo != null)
            {
                this.curITypeInfo.ReleaseTypeAttr(this.curTypeAttrPtr);
            }

            this.ppTLB.GetTypeInfo(this.curTypeInfo, out this.curITypeInfo);

            this.curITypeInfo.GetTypeAttr(out this.curTypeAttrPtr);
            this.curTypeAttr = Marshal.PtrToStructure<TYPEATTR>(curTypeAttrPtr);

            return true;
        }

        public bool IsTypeKind(int kind)
        {
            return this.curTypeAttr.typekind == (TYPEKIND)kind;
        }

        public bool IsTypeEnum() => IsTypeKind((int)TYPEKIND.TKIND_ENUM);
        public bool IsTypeRecord() => IsTypeKind((int)TYPEKIND.TKIND_RECORD);
        public bool IsTypeModule() => IsTypeKind((int)TYPEKIND.TKIND_MODULE);
        public bool IsTypeInterface() => IsTypeKind((int)TYPEKIND.TKIND_INTERFACE);
        public bool IsTypeDispatch() => IsTypeKind((int)TYPEKIND.TKIND_DISPATCH);
        public bool IsTypeCoClass() => IsTypeKind((int)TYPEKIND.TKIND_COCLASS);
        public bool IsTypeAlias() => IsTypeKind((int)TYPEKIND.TKIND_ALIAS);
        public bool IsTypeUnion() => IsTypeKind((int)TYPEKIND.TKIND_UNION);
        public bool IsTypeMax() => IsTypeKind((int)TYPEKIND.TKIND_MAX);

        public string LibDocumentation()
        {
            if (this.ppTLB == null) return "No Type Library open!";

            this.ppTLB.GetDocumentation(
              -1,
              out string name,
              out string doc,
              out int ctx,
              out string helpfile  // Help File
              );

            return name + ": " + doc;
        }

        public string TypeDocumentation()
        {
            return TypeDocumentation(this.curITypeInfo);
        }

        private string TypeDocumentation(ITypeInfo typeInfo)
        {
            typeInfo.GetDocumentation(
              -1, // MEMBERID_NIL, //idx,
              out string name,
              out string doc,
              out int ctx,
              out string helpFile  // Help File
              );

            return name;
        }

        public enum INVOKEKIND
        {
            func = 0x01,
            get = 0x02,
            put = 0x04,
            putref = 0x08
        };

        public INVOKEKIND InvokeKind()
        {
            return (INVOKEKIND)(this.curFuncDesc.invkind);
        }

        public enum VARIABLEKIND
        {
            instance = 0, // VAR_PERINSTANCE,
            static_ = 1, // VAR_STATIC,
            const_ = 2, // VAR_CONST,
            dispatch = 3 // VAR_DISPATCH
        };

        public VARIABLEKIND VariableKind()
        {
            return (VARIABLEKIND)(this.curVarDesc.varkind);
        }

        public bool NextFunction()
        {
            this.curFunc++;
            this.curFuncParam = -1;

            if (this.curFunc >= this.curTypeAttr.cFuncs) return false;

            if (this.curFuncDescPtr != null)
            {
                this.curITypeInfo.ReleaseFuncDesc(curFuncDescPtr);
            }

            this.curITypeInfo.GetFuncDesc(this.curFunc, out this.curFuncDescPtr);
            this.curFuncDesc = Marshal.PtrToStructure<FUNCDESC>(curFuncDescPtr);
            GetFuncNames();
            this.elemDescArray = new ELEMDESC[curFuncDesc.cParams];
            for (int i = 0; i<curFuncDesc.cParams; i++)
            {
                ELEMDESC ed = Marshal.PtrToStructure<ELEMDESC>(this.curFuncDesc.lprgelemdescParam + i * Marshal.SizeOf<ELEMDESC>());
                this.elemDescArray[i] = ed;

            }
            return true;
        }

        public string FunctionName()
        {
            return this.funcNames[0];
        }

        public string ConstValue()
        {
            if (VariableKind() != VARIABLEKIND.const_)
            {
                return "VariableKind must be const_";
            }
            object o = Marshal.GetObjectForNativeVariant(this.curVarDesc.desc.lpvarValue);

            return o.ToString(); //  v.ValueAsString();
        }

        //public bool HasFunctionTypeFlag(TYPEFLAG tf)
        //{
        //    return curImplTypeFlags_ & static_cast<int>(tf);
        //}

        string UserdefinedType(int hrt)
        {
            string tp = "";

            this.curITypeInfo.GetRefTypeInfo(hrt, out ITypeInfo itpi);
            string ref_type = TypeDocumentation(itpi);
            tp += ref_type;

            return tp;
        }

        TypeInfo TypeInfoFactory(ELEMDESC ed)
        {
            TYPEDESC td = ed.tdesc;
            return RecTypeInfoFactory(td);
        }

        TypeInfo RecTypeInfoFactory(TYPEDESC td)
        {
            TypeInfo ti = null;
            switch ((VarEnum)td.vt) {
                case VarEnum.VT_EMPTY:
                    ti = new UnknownTypeInfo("VT_EMPTY");
                    break;
                case VarEnum.VT_NULL:
                    ti = new UnknownTypeInfo("VT_NULL");
                    break;
                case VarEnum.VT_I2:
                case VarEnum.VT_I4:
                case VarEnum.VT_R4:
                case VarEnum.VT_R8:
                case VarEnum.VT_CY:
                case VarEnum.VT_DATE:
                case VarEnum.VT_BSTR:
                case VarEnum.VT_DISPATCH:
                case VarEnum.VT_ERROR:
                case VarEnum.VT_BOOL:
                    ti = new ScalarTypeInfo((VarEnum)td.vt);
                    break;
                case VarEnum.VT_VARIANT:
                    ti = new UnknownTypeInfo("VT_VARIANT");
                    break;
                case VarEnum.VT_UNKNOWN:
                    ti = new UnknownTypeInfo("VT_UNKNOWN");
                    break;
                case VarEnum.VT_DECIMAL:
                case VarEnum.VT_I1:
                case VarEnum.VT_UI1:
                case VarEnum.VT_UI2:
                case VarEnum.VT_UI4:
                case VarEnum.VT_I8:
                case VarEnum.VT_UI8:
                case VarEnum.VT_INT:
                case VarEnum.VT_UINT:
                case VarEnum.VT_VOID:
                case VarEnum.VT_HRESULT:
                    ti = new ScalarTypeInfo((VarEnum)td.vt);
                    break;
                case VarEnum.VT_PTR:
                    {
                        TYPEDESC ptrTd = Marshal.PtrToStructure<TYPEDESC>(td.lpValue);
                        ti = new PointerTypeInfo(RecTypeInfoFactory(ptrTd));
                    }
                    break;
                case VarEnum.VT_SAFEARRAY:
                    {
                        TYPEDESC saTd = Marshal.PtrToStructure<TYPEDESC>(td.lpValue);
                        ti = new SafeArrayTypeInfo(RecTypeInfoFactory(saTd));
                    }
                    break;
                case VarEnum.VT_CARRAY:
                    ti = new UnknownTypeInfo("VT_CARRAY");
                    break;
                case VarEnum.VT_USERDEFINED:
                    {                        
                        ITypeInfo refTi = null;
                        Int32 hreftype = td.lpValue.ToInt32();
                        curITypeInfo.GetRefTypeInfo(hreftype, out refTi);
                        refTi.GetDocumentation(-1, out string name, out _, out _, out _);
                        ti = new UserTypeInfo(name); //  UserdefinedType(refTd));
                    }
                    break;

                case VarEnum.VT_LPSTR:
                case VarEnum.VT_LPWSTR:
                case VarEnum.VT_RECORD:
                case VarEnum.VT_FILETIME:
                case VarEnum.VT_BLOB:
                case VarEnum.VT_STREAM:
                case VarEnum.VT_STORAGE:
                case VarEnum.VT_STREAMED_OBJECT:
                case VarEnum.VT_STORED_OBJECT:
                case VarEnum.VT_BLOB_OBJECT:
                case VarEnum.VT_CF:
                case VarEnum.VT_CLSID:
                case VarEnum.VT_VECTOR:
                case VarEnum.VT_ARRAY:
                case VarEnum.VT_BYREF:
                    ti = new UnknownTypeInfo("Other unhandled type: " + td.vt.ToString());
                    break;
                default:
                    ti = new UnknownTypeInfo("Unknown type: " + td.vt.ToString());
                    break;
            }
            return ti;
        }

        public string ParameterType()
        {
            TypeInfo ti = ParameterTypeInfo();
            return ti.ToString();
        }

        public TypeInfo ParameterTypeInfo()
        {
            ELEMDESC ed = this.elemDescArray[this.curFuncParam];
            return TypeInfoFactory(ed);
        }

        public string VariableType()
        {
            TypeInfo ti = VariableTypeInfo();
            return ti.ToString();
        }

        TypeInfo VariableTypeInfo()
        {
            ELEMDESC ed = this.curVarDesc.elemdescVar;
            return TypeInfoFactory(ed);
        }

        public string ReturnType()
        {
            TypeInfo ti = ReturnTypeInfo();
            return ti.ToString();
        }

        TypeInfo ReturnTypeInfo()
        {
            ELEMDESC ed = this.curFuncDesc.elemdescFunc;
            return TypeInfoFactory(ed);
        }

        public bool NextVariable()
        {
            this.curVar++;
            // TODO: curVar_->cVars accessed through NofVariables().
            if (this.curVar >= this.curTypeAttr.cVars) return false;

            if (this.curVarDescPtr != IntPtr.Zero)
            {
                this.curITypeInfo.ReleaseVarDesc(this.curVarDescPtr);
            }

            this.curITypeInfo.GetVarDesc(this.curVar, out this.curVarDescPtr);
            this.curVarDesc = Marshal.PtrToStructure<VARDESC>(curVarDescPtr);
            return true;
        }

        public bool NextParameter()
        {
            this.curFuncParam++;
            if (curFuncParam >= NofParameters()) return false;

            return true;
        }

        public string VariableName()
        {
            string[] varName = new string[1];
            this.curITypeInfo.GetNames(this.curVarDesc.memid, varName, 1, out int dummy);
            return varName[0];
        }

        public string ParameterName()
        {
            if (1 + this.curFuncParam >= this.funcNamesCount) return "<nameless>";
            string paramName = this.funcNames[1 + this.curFuncParam];
            return paramName;
        }

        // http://doc.ddart.net/msdn/header/include/oaidl.h.html
        //  return ParameterIsHasX(PARAMFLAG_NONE);
        public bool ParameterIsIn()
        {
            return ParameterIsHasX(PARAMFLAG.PARAMFLAG_FIN);
        }
        public bool ParameterIsOut()
        {
            return ParameterIsHasX(PARAMFLAG.PARAMFLAG_FOUT);
        }
        bool ParameterIsFLCID()
        {
            return ParameterIsHasX(PARAMFLAG.PARAMFLAG_FLCID);
        }
        bool ParameterIsReturnValue()
        {
            return ParameterIsHasX(PARAMFLAG.PARAMFLAG_FRETVAL);
        }
        bool ParameterIsOptional()
        {
            return ParameterIsHasX(PARAMFLAG.PARAMFLAG_FOPT);
        }
        bool ParameterHasDefault()
        {
            return ParameterIsHasX(PARAMFLAG.PARAMFLAG_FHASDEFAULT);
        }
        bool ParameterHasCustumData()
        {
            return ParameterIsHasX((PARAMFLAG)0x40 /*PARAMFLAG_FHASCUSTDATA*/);
        }
        bool ParameterIsHasX(PARAMFLAG flag)
        {
            ELEMDESC e = this.elemDescArray[this.curFuncParam];
            return (e.desc.paramdesc.wParamFlags & flag) > 0;
        }

        void GetFuncNames()
        {
            int nof_names = 1 + NofParameters();
            this.funcNames = new string[nof_names];

            this.curITypeInfo.GetNames(this.curFuncDesc.memid, this.funcNames, nof_names, out this.funcNamesCount);
        }

        public int NofVariables()
        {
            return this.curTypeAttr.cVars;
        }

        public int NofFunctions()
        {
            return this.curTypeAttr.cFuncs;
        }

        public int NofTypeInfos()
        {
            return this.typeInfoCount;
        }

        public int NofParameters()
        {
            return this.curFuncDesc.cParams;
        }

        public int NofOptionalParameters()
        {
            return this.curFuncDesc.cParamsOpt;
        }

        public Dictionary<string, Dictionary<string, int>> EnumDictionary()
        {
            var rv = new Dictionary<string, Dictionary<string, int>>();
            while (this.NextTypeInfo())
            {
                string typeDoc = this.TypeDocumentation();
                if (this.IsTypeEnum())
                {
                    var en = new Dictionary<string, int>();
                    rv[typeDoc] = en;
                    while (this.NextVariable())
                    {
                        TypeLibEnumerator.VARIABLEKIND vk = this.VariableKind();
                        if (vk == TypeLibEnumerator.VARIABLEKIND.const_)
                        {
                            if(int.TryParse(this.ConstValue(), out int enumValue))
                            {
                                en[this.VariableName()] = enumValue;
                            }
                        }
                    }
                }
            }
            return rv;
        }

        private int typeInfoCount = 0;
        private int curTypeInfo = -1;
        private ITypeLib ppTLB = null;

        private int curFunc = -1;
        private int curVar = -1;
        private int curFuncParam = -1;

        private ITypeInfo curITypeInfo = null;
        private IntPtr curTypeAttrPtr = IntPtr.Zero;
        private TYPEATTR curTypeAttr;

        private int funcNamesCount = 0;
        private FUNCDESC curFuncDesc;
        private IntPtr curFuncDescPtr = IntPtr.Zero;
        private VARDESC curVarDesc;
        private IntPtr curVarDescPtr = IntPtr.Zero;
        private string[] funcNames;
        private int paramCount;
        ELEMDESC[] elemDescArray;
    }

    static class EnumContainer
    {
        public static System.Dynamic.ExpandoObject createEnumContainer(
            Dictionary<string, Dictionary<string, int>> enumDict)
        {
            dynamic exoContainer = new System.Dynamic.ExpandoObject();
            foreach (var edd in enumDict)
            {
                dynamic exoType = new System.Dynamic.ExpandoObject();
                foreach(var ed in edd.Value)
                {
                    ((IDictionary<String, Object>)exoType).Add(ed.Key, ed.Value);
                }               
                ((IDictionary<String, Object>)exoContainer).Add(edd.Key, exoType);
            }
            return exoContainer;
        }
    }

}
