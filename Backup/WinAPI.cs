using System;
using System.Runtime.InteropServices;

class WinAPI {

	public const int SW_HIDE=0;

	[DllImport("User32.dll", EntryPoint="SetParent")]
	public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

}