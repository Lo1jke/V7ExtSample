using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace V7ExtSample
{
	[ComVisible(true), Guid("0339E962-6744-4844-9408-1A795828CCBF"), ProgId("AddIn.NetComponentSampleCS")]
	public class NetComponentSampleCS : AddInLib.IInitDone, AddInLib.ILanguageExtender 
	{
		const string c_AddinName  = "NetComponentSampleCS";
							  
		#region "����������"
		string strMessageBoxIcon;
		string strMessageBoxButtons;
		#endregion
															   
		public NetComponentSampleCS()
		{

		}
		#region "IInitDone implementation"
		public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
		{
			V7Data.V7Object = pConnection;

			this.strMessageBoxButtons = "OK";
			this.strMessageBoxIcon = "None";

		}

		public void Done()
		{
			//MessageBox.Show("*Done", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}

		public void GetInfo(ref object pInfo)
		{
			((System.Array)pInfo).SetValue("2000", 0);
		}
		#endregion

		public void RegisterExtensionAs(ref string bstrExtensionName)
		{            
			bstrExtensionName = c_AddinName;
		}

		
		#region "��������"
		enum Props
		{   //�������� �������������� ������� ����� ������� ����������
			propMessageBoxIcon = 0,  //����������� ��� MessageBox'�
			propMessageBoxButtons = 1, //������ ��� MessageBox'a
			LastProp = 2
		}

		public void GetNProps(ref int plProps)
		{	//����� 1� �������� ���������� ��������� �� �� �������
			plProps = (int)Props.LastProp;
		}
        
		public void FindProp(string bstrPropName, ref int plPropNum)
		{	//����� 1� ���� �������� ������������� �������� �� ��� ���������� �����
			switch(bstrPropName)
			{
				case "MessageBoxIcon":
				case "�������������������������":
					plPropNum = (int)Props.propMessageBoxIcon;
					break;
				case "MessageBoxButtons":
				case "��������������������":
					plPropNum = (int)Props.propMessageBoxButtons;
					break;
				default:
					plPropNum = -1;
					break;
			}
		}
		
		public void GetPropName(int lPropNum, int lPropAlias, ref string pbstrPropName)
		{	//����� 1� (������������) ������ ��� �������� �� ��� ��������������. lPropAlias - ����� ����������
			pbstrPropName = "";
		}

		public void GetPropVal(int lPropNum, ref object pvarPropVal)
		{	//����� 1� ������ �������� ������� 
			pvarPropVal = null;
			switch(lPropNum)
			{
				case (int)Props.propMessageBoxIcon:
					pvarPropVal = strMessageBoxIcon;
					break;
				case (int)Props.propMessageBoxButtons:
					pvarPropVal = strMessageBoxButtons;
					break;
			}
		}
		
		public void SetPropVal(int lPropNum, ref object varPropVal)
		{	//����� 1� �������� �������� ������� 

			switch(lPropNum)
			{
				case (int)Props.propMessageBoxIcon:
					strMessageBoxIcon = (string)varPropVal;
					break;
				case (int)Props.propMessageBoxButtons:
					strMessageBoxButtons = (string)varPropVal;
					break;

			}
		}
		
		public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
		{	//����� 1� ������, ����� �������� �������� ��� ������
			pboolPropRead = true; // ��� �������� �������� ��� ������
		}
		
		public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
		{	//����� 1� ������, ����� �������� �������� ��� ������
			pboolPropWrite = true; // ��� �������� �������� ��� ������
		}
		#endregion
	
		#region "������"

		enum Methods
		{	//�������� �������������� ������� (�������� ��� �������) ����� ������� ����������
			methMessageBoxShow = 0, //��������������, �� � ������������ �������� ����������� � ��������� ����
			methExternalEvent = 1, //���������� ������� ������� (1� ������������� ��� � ��������� ������������������������())
			methShowErrorLog = 2, //���������� ��������� ��������� 
			methStatusLine = 3, //���������� ��������� � ������ ���������
			LastMethod = 4,
		}

		public void GetNMethods(ref int plMethods)
		{	//����� 1� �������� ���������� ��������� �� �� �������
			plMethods = (int)Methods.LastMethod;
		}
		
		public void FindMethod(string bstrMethodName, ref int plMethodNum)
		{	//����� 1� �������� �������� ������������� ������ (��������� ��� �������) �� ����� (��������) ��������� ��� �������
			plMethodNum = -1;
			switch(bstrMethodName)
			{
				case "MessageBoxShow":
				case "��������������":
					plMethodNum = (int)Methods.methMessageBoxShow;
					break;
				case "ExternalEvent":
				case "��������������":
					plMethodNum = (int)Methods.methExternalEvent;
					break;
				case "ShowErrorLog":
				case "��������":
					plMethodNum = (int)Methods.methShowErrorLog;
					break;
				case "StatusLine":
				case"���������":
					plMethodNum = (int)Methods.methStatusLine;
					break;

			}
		}
		
		public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
		{	//����� 1� (������������) �������� ��� ������ �� ��� ��������������. lMethodAlias - ����� ��������.
			pbstrMethodName = "";
		}

		public void GetNParams(int lMethodNum, ref int plParams)
		{	//����� 1� �������� ���������� ���������� � ������ (��������� ��� �������)
			switch(lMethodNum)
			{
				case (int)Methods.methMessageBoxShow:
					plParams = 2;
					break;
				case (int)Methods.methExternalEvent:
					plParams = 3;
					break;
				case (int)Methods.methShowErrorLog:
					plParams = 1;
					break;
				case (int)Methods.methStatusLine:
					plParams = 1;
					break;
			}
		}
		
		public void GetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
		{	//����� 1� �������� �������� ���������� ��������� ��� ������� �� ���������
			pvarParamDefValue = null; //��� �������� �� ���������
		}

		public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
		{	//����� 1� ������, ���������� �� ����� �������� (�.�. �������� ���������� ��� ��������)
			pboolRetValue = true;  //��� ������ � ��� ����� ��������� (�.�. ����� ���������� ��������). 
		}

		public void CallAsProc(int lMethodNum, ref System.Array paParams)
		{	//����� ������� ���������� ��������� ��� ��������. � �������� � ��� ���.
		}

		public void CallAsFunc(int lMethodNum, ref object pvarRetValue, ref System.Array paParams)
		{	//����� ������� ���������� ��������� ��� �������.
			MessageBoxIcon icon;
			pvarRetValue = 0; //������������ �������� ������ ��� 1�
			
			switch(lMethodNum) //���������� ����� ������
			{
				case (int)Methods.methMessageBoxShow: //��������� ����� MessageBoxShow ������� ����������
				{
					icon = MessageBoxIcon.None;
					//��������������� ��������� �������� ������ � MessageBoxIcon.���
					switch(strMessageBoxIcon)
					{
						case "Asterisk":
							icon = MessageBoxIcon.Asterisk;
							break;
						case "Error":
							icon = MessageBoxIcon.Error;
							break;
						case "Exclamation":
							icon = MessageBoxIcon.Exclamation;
							break;
						case "Hand":
							icon = MessageBoxIcon.Hand;
							break;
						case "Information":
							icon = MessageBoxIcon.Information;
							break;
						case "None":
							icon = MessageBoxIcon.None;
							break;
						case "Question":
							icon = MessageBoxIcon.Question;
							break;
						case "Stop":
							icon = MessageBoxIcon.Stop;
							break;
						case "Warning":
							icon = MessageBoxIcon.Warning;
							break;
					}


					//��������������� ��������� �������� ������ � MessageBoxButtons.���
					MessageBoxButtons butt = MessageBoxButtons.OK;
					switch(strMessageBoxButtons)
					{
						case "AbortRetryIgnore":
							butt = MessageBoxButtons.AbortRetryIgnore;
							break;
						case "OK":
							butt = MessageBoxButtons.OK;
							break;
						case "OKCancel":
							butt = MessageBoxButtons.OKCancel;
							break;
						case "RetryCancel":
							butt = MessageBoxButtons.RetryCancel;
							break;
						case "YesNo":
							butt = MessageBoxButtons.YesNo;
							break;
						case "YesNoCancel":
							butt = MessageBoxButtons.YesNoCancel;
							break;
					}
                    
					DialogResult res;
					string strMessageBoxText; 
					string strMessageBoxHeader;
					string strDialogResult;
                    
					
					strMessageBoxText = (string)paParams.GetValue(0); //�������� ������ �������� ����� ������� - ����� ��������������
					strMessageBoxHeader = (string)paParams.GetValue(1);//�������� ������ �������� ����� ������� - ��������� ��������������
					
					//���������� ���������� ���� MessageBox.Show 
					res = MessageBox.Show(strMessageBoxText, strMessageBoxHeader, butt, icon);
					//��������������� ��������� �� DialogResult.��� � ��������� ������
					switch(res)
					{
						case DialogResult.Abort:
							strDialogResult = "Abort";
							break;
						case DialogResult.Cancel:
							strDialogResult = "Cancel";
							break;
						case DialogResult.Ignore:
							strDialogResult = "Ignore";
							break;
						case DialogResult.No:
							strDialogResult = "No";
							break;
						case DialogResult.None:
							strDialogResult = "None";
							break;
						case DialogResult.OK:
							strDialogResult = "OK";
							break;
						case DialogResult.Retry:
							strDialogResult = "Retry";
							break;
						case DialogResult.Yes:
							strDialogResult = "Yes";
							break;
						default:
							strDialogResult = "";
							break;
					}
					
					pvarRetValue = strDialogResult; //������������ ��������
					
					//���������� ��� �������� � �������� ���������
					strMessageBoxButtons = "OK";
					strMessageBoxIcon = "None";
					break;
				} // ����� ������ MessageBosIcon
					//////////////////////////////////////////////////////////
					
				case (int)Methods.methExternalEvent:  //��������� ����� ��� ��������� �������� �������
				{
					string s1;
					string s2;
					string s3;
					s1 = (string)paParams.GetValue(0);
					s2 = (string)paParams.GetValue(1);
					s3 = (string)paParams.GetValue(2);
					V7Data.AsyncEvent.ExternalEvent(s1, s2, s3);
					break;
				} // ����� ������ ExternalEvent
					//////////////////////////////////////////////////////////
                
				case (int)Methods.methShowErrorLog:  //��������� ����� ��� ������ ��������� �� ������
				{
					string s1;
					s1 = (string)paParams.GetValue(0);
                    
					AddInLib.ExcepInfo ei = new AddInLib.ExcepInfo();
					ei.wCode = 1006; //��� �����������
					ei.bstrDescription = s1; 
					ei.bstrSource = c_AddinName;
                    
					V7Data.ErrorLog.AddError("", ei);
					throw new System.Exception("An exception has occurred.");
					break;
				} // ����� ������ ShowErrorLog
					//////////////////////////////////////////////////////////
				
				case (int)Methods.methStatusLine: //��������� �������� ����� ��� ��������� ������ ���������
				{
					string s1;
					s1 = (string)paParams.GetValue(0);
					V7Data.StatusLine.SetStatusLine(s1);
					System.Threading.Thread.Sleep(1000); //������ ����� 1 ���
					V7Data.StatusLine.ResetStatusLine();
					break;
				}
			}
		}
		#endregion
	}
}
