using System;
using System.Runtime.InteropServices;

namespace V7ExtSample.AddInLib
{
	[Guid("ab634005-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IStatusLine
	{
		void SetStatusLine(string bstrStatusLine);
		void ResetStatusLine();
	}
}