using System;
using System.Runtime.InteropServices;

namespace V7ExtSample.AddInLib
{
	[Guid("ab634004-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IAsyncEvent
	{
		void SetEventBufferDepth(int lDepth);
		void GetEventBufferDepth(ref int plDepth);
		void ExternalEvent(string bstrSource, string bstrMessage, string bstrData);
		void CleanBuffer();
	}
}
