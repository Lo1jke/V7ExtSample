using System;
using System.Runtime.InteropServices;

namespace V7ExtSample.AddInLib
{
	[Guid("3127CA40-446E-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IErrorLog
	{
		void AddError(string pszPropName, ExcepInfo pExepInfo);
	}
	//----------------------------------------------------------
	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode, Pack=8)]
	public struct ExcepInfo
	{
		public short wCode;
		public short wReserved;
		[MarshalAs(UnmanagedType.BStr)] public string bstrSource;
        [MarshalAs(UnmanagedType.BStr)] public string bstrDescription;
		[MarshalAs(UnmanagedType.BStr)] public string bstrHelpFile;
		public int dwHelpContext;
		public System.IntPtr pvReserved;
		public System.IntPtr pfnDereffered;
		public int scode;
	}
}