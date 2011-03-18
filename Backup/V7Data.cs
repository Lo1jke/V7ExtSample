using System;

namespace V7ExtSample
{
	internal class V7Data
	{

		public static object V7Object
		{
			get
			{
				return m_V7Object;
			}
			set
			{
				m_V7Object = value;
				// Вызываем неявно QueryInterface
				m_ErrorInfo = (AddInLib.IErrorLog)value;
				m_AsyncEvent = (AddInLib.IAsyncEvent)value;
				m_StatusLine = (AddInLib.IStatusLine)value;
			}
		}

		public static AddInLib.IErrorLog ErrorLog
		{
			get
			{
				return m_ErrorInfo;
			}
		}

		public static AddInLib.IAsyncEvent AsyncEvent
		{
			get
			{
				return m_AsyncEvent;
			}
		}

		public static AddInLib.IStatusLine StatusLine
		{
			get
			{
				return m_StatusLine;
			}
		}


		private static object m_V7Object;
		private static AddInLib.IErrorLog m_ErrorInfo;
		private static AddInLib.IAsyncEvent m_AsyncEvent;
		private static AddInLib.IStatusLine m_StatusLine;

	}
}
