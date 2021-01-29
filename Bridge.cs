using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TestWebViewWPF
{



    [ComVisible(true)]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        const int DISP_VALUE = 0;
        const int LOCALE_USER_DEFAULT = 0x400;

        int GetTypeInfoCount(out uint pctinfo);
        int GetTypeInfo(uint iTInfo, int lcid, out IntPtr info);

        int GetIDsOfNames(
            ref Guid iid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] names,
            int cNames,
            int lcid,
            [Out][MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2)] int[] rgDispId);

        int Invoke(
            int dispId,
            ref Guid riid,
            int lcid,
            System.Runtime.InteropServices.ComTypes.INVOKEKIND wFlags,
            ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
            [Out][MarshalAs(UnmanagedType.SafeArray)] out object[] result,
            ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo,
            IntPtr puArgErr);
    }



    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class Bridge
    {

        [DllImport("oleaut32.dll")]
        static extern void VariantClear(IntPtr pVariant);

        public string Name { get; set; } = "WebView2";

        public string NativeFunction(string param)
        {
            System.Console.WriteLine("Native function with param " + param);
            return "This string is returned from C# using webview2";
        }

        public bool NativeFunctionWithCallBack(IDispatch callback)
        {
            System.Console.WriteLine("NativeFunction with callback");
            if (callback != null)
            {

                object[] result = null;
                System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams = new System.Runtime.InteropServices.ComTypes.DISPPARAMS();
                pDispParams.cArgs = 2;
                Guid guid = Guid.Empty;

                int variantSize = Marshal.SizeOf<Variant>();

                // Marshal a value to a variant
                object[] values = new object[] { "String param from webview2", 2 };
                IntPtr pVariantArgArray = Marshal.AllocCoTaskMem(variantSize * values.Length);
                for (int i = 0; i < values.Length; ++i)
                {
                    // !! The arguments should be in REVERSED order!!
                    int actualIndex = (values.Length - i - 1);
                    Marshal.GetNativeVariantForObject(values[i], pVariantArgArray + actualIndex * variantSize);
                }
                pDispParams.rgvarg = pVariantArgArray;


                System.Runtime.InteropServices.ComTypes.EXCEPINFO execpt_info = default(System.Runtime.InteropServices.ComTypes.EXCEPINFO);

                int r = callback.Invoke(IDispatch.DISP_VALUE, ref guid, IDispatch.LOCALE_USER_DEFAULT, System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC, ref pDispParams, out result, ref execpt_info, IntPtr.Zero);

               
                if (pVariantArgArray != IntPtr.Zero)
                {
                    for (int i = 0; i < values.Length; ++i)
                    {
                        VariantClear(pVariantArgArray + i * variantSize);
                    }

                    Marshal.FreeCoTaskMem(pVariantArgArray);
                }

                if (r == 0) //S_OK 
                {
                    return true;
                }

            }

            return false;
        }
    }

    #region Variant type
    [StructLayout(LayoutKind.Explicit)]
    public struct Variant
    {
        // Most of the data types in the Variant are carried in _typeUnion
        [FieldOffset(0)]
        internal TypeUnion _typeUnion;

        // Decimal is the largest data type and it needs to use the space that is normally unused in TypeUnion._wReserved1, etc.
        // Hence, it is declared to completely overlap with TypeUnion. A Decimal does not use the first two bytes, and so
        // TypeUnion._vt can still be used to encode the type.
        [FieldOffset(0)]
        internal Decimal _decimal;

        [StructLayout(LayoutKind.Explicit)]
        internal struct TypeUnion
        {
            [FieldOffset(0)]
            internal ushort _vt;
            [FieldOffset(2)]
            internal ushort _wReserved1;
            [FieldOffset(4)]
            internal ushort _wReserved2;
            [FieldOffset(6)]
            internal ushort _wReserved3;
            [FieldOffset(8)]
            internal UnionTypes _unionTypes;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct Record
        {
            internal IntPtr _record;
            internal IntPtr _recordInfo;
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct UnionTypes
        {
            [FieldOffset(0)]
            internal SByte _i1;
            [FieldOffset(0)]
            internal Int16 _i2;
            [FieldOffset(0)]
            internal Int32 _i4;
            [FieldOffset(0)]
            internal Int64 _i8;
            [FieldOffset(0)]
            internal Byte _ui1;
            [FieldOffset(0)]
            internal UInt16 _ui2;
            [FieldOffset(0)]
            internal UInt32 _ui4;
            [FieldOffset(0)]
            internal UInt64 _ui8;
            [FieldOffset(0)]
            internal Int32 _int;
            [FieldOffset(0)]
            internal UInt32 _uint;
            [FieldOffset(0)]
            internal Int16 _bool;
            [FieldOffset(0)]
            internal Int32 _error;
            [FieldOffset(0)]
            internal Single _r4;
            [FieldOffset(0)]
            internal Double _r8;
            [FieldOffset(0)]
            internal Int64 _cy;
            [FieldOffset(0)]
            internal double _date;
            [FieldOffset(0)]
            internal IntPtr _bstr;
            [FieldOffset(0)]
            internal IntPtr _unknown;
            [FieldOffset(0)]
            internal IntPtr _dispatch;
            [FieldOffset(0)]
            internal IntPtr _pvarVal;
            [FieldOffset(0)]
            internal IntPtr _byref;
            [FieldOffset(0)]
            internal Record _record;
        }
    }

    #endregion
}
