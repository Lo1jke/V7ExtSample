﻿using System;
using System.Runtime.InteropServices;

namespace V7ExtSample.AddInLib
{
	[Guid("AB634001-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IInitDone
	{
		void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection);
		void Done();
		void GetInfo(ref object pInfo);
	}
}