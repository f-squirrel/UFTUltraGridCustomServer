using System;
using Mercury.QTP.CustomServer;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Reflection;
using System.Collections.Generic;

namespace UFTUltraGridCustomServer
{
    [ReplayInterface]
    public interface IUltraGridUFTCustomServerReplay
    {
        #region Wizard generated sample code (commented)
        //		void  CustomMouseDown(int X, int Y);
        #endregion

        Object GetNAProperty(String propertyPath);
    }
    /// <summary>
    /// Summary description for UltraGridUFTCustomReplayRecordServer.
    /// </summary>
    public class UltraGridUFTCustomReplayRecordServer :
        CustomServerBase,
        IUltraGridUFTCustomServerReplay
    {
        // Do not call Base class methods or properties in the constructor.
        // The services are not initialized.
        public UltraGridUFTCustomReplayRecordServer()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        #region IRecord override Methods
        #region Wizard generated sample code (commented)
        /*		/// <summary>
		/// To change Window messages filter, implement this method.
		/// The default implementation is to get only the 
        /// Control's window messages.
		/// </summary>
		public override WND_MsgFilter GetWndMessageFilter()
		{
			return WND_MsgFilter.WND_MSGS;
		}

		/// <summary>
		/// To catch window messages, implement this method.
		/// This method is called only if the CustomServer is running
		/// in the QuickTest context.
		/// </summary>
		public override RecordStatus OnMessage(ref Message tMsg)
		{
			// TODO:  Add OnMessage implementation.
			return RecordStatus.RECORD_HANDLED;
		}
*/
        #endregion
        /// <summary>
        /// To extend the Record process, add Events handlers
        /// to listen to the custom control's events.
        /// </summary>
        public override void InitEventListener()
        {
            #region Wizard generated sample code (commented)
            /*			// You can add as many handlers as you need.
			// For example, to add an OnMouseDown handler, 
            // first create the Delegate:
			Delegate  e = new System.Windows.Forms.MouseEventHandler(this.OnMouseDown);

            // Then, add the event handler as the first handler of the event.
			// The first argument is the name of the event for which to listen.
			// This must be an event that the control supports. You can
			// use the .NET Spy to obtain the list of events supported by the control.
			// The second argument is the event handler delegate. 

			AddHandler("MouseDown", e);
*/
            #endregion
        }

        /// <summary>
        /// Called by QuickTest to release the handlers
        /// added in the InitEventListener method.
        /// Only handlers added using QuickTest methods are released 
        /// by the QuickTest infrastructure. If you use standard C# syntax,
        /// you must release the handlers in your code at the end 
        /// of the Record process.
        /// </summary>
        public override void ReleaseEventListener()
        {
        }

        #endregion


        #region Record events handlers
        #region Wizard generated sample code (commented)
        /*		public void OnMouseDown(object sender, System.Windows.Forms.MouseEventArgs  e)
		{
			// This example shows how to create a script line in QuickTest
			// when a MouseDown event is encountered during recording.
			if(e.Button == System.Windows.Forms.MouseButtons.Left)
			{			
				RecordFunction( "CustomMouseDown", RecordingMode.RECORD_SEND_LINE, e.X, e.Y);
			}
		}
*/
        #endregion
        #endregion


        #region Replay interface implementation
        #region Wizard generated sample code (commented)
        /*		public void CustomMouseDown(int X, int Y)
		{
			MouseClick(X, Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
		}
*/
 
        private List<ObjectPropertyDescriptor> GetListOfProperties(string propertyPath)
        {
            if (propertyPath.Length == 0)
            {
                throw new ArgumentException("Empty property path");
            }

            List<ObjectPropertyDescriptor> list = new List<ObjectPropertyDescriptor>();
            int startOfPropertyName = 0;
            for (int i = 0; i < propertyPath.Length; i++)
            {
                char ch = propertyPath[i];
                while (i < propertyPath.Length && (Char.IsLetter(propertyPath[i]) || Char.IsNumber(propertyPath[i])))
                {
                    ++i;
                }
                string propertyName = propertyPath.Substring(startOfPropertyName, i - startOfPropertyName);
                startOfPropertyName = i;

                var descriptor = new ObjectPropertyDescriptor(propertyName);

                if (i < propertyPath.Length && propertyPath[i] == '[')
                {
                    descriptor.IsIndexed = true;
                    int indexStart = ++i;
                    while (i < propertyPath.Length && propertyPath[i] != ']')
                    {
                        ++i;
                    }
                    string index = propertyPath.Substring(indexStart, i - indexStart);
                    descriptor.Index = Int32.Parse(index);
                    ++i;
                }
                list.Add(descriptor);

                if (i >= propertyPath.Length)
                {
                    break;
                }

                if (propertyPath[i] == '.')
                {
                    startOfPropertyName = ++i;
                    continue;
                }
                else
                {
                    throw new Exception("Incorrect property path: mot (.) after property name or index");
                }
            }

            return list;
        }

        public Object GetNAProperty(String propertyPath)
        {

            var result = GetNAProperty(propertyPath, (Infragistics.Win.UltraWinGrid.UltraGrid)SourceControl);
            ReplayReportStep("GetNAProperty", EventStatus.EVENTSTATUS_GENERAL, result);
            return result;
        }

        public Object GetNAProperty(String propertyPath, object obj)
        {
            if (propertyPath.Length == 0)
            {
                ReplayThrowError("Property name is empty!");
            }

            List<ObjectPropertyDescriptor> descriptorLIst = GetListOfProperties(propertyPath);

            object currentValue = obj;
            for (int i = 0; i < descriptorLIst.Count; i++)
            {
                ObjectPropertyDescriptor descr = descriptorLIst[i];

                PropertyInfo property = currentValue.GetType().GetProperty(descr.Name);
                if (property == null)
                {
                    ReplayThrowError("Cannot find property: " + descr.Name);
                    break;
                }

                ParameterInfo[] parms = property.GetIndexParameters();
                object value = null;
                if (parms.Length == 0)
                {
                    value = property.GetValue(currentValue, null);
                    if (descr.IsIndexed)
                    {
                        Type type = value.GetType();
                        if (type.IsArray)
                        {
                            value = type.GetMethod("Get").Invoke(value, new object[] { descr.Index });
                        }
                        else // It looks like Collection
                        {
                            property = type.GetProperty("Item");
                            if (property == null)
                            {
                                ReplayThrowError("Cannot find property: \"Item\" for " + descr.Name);
                                break;
                            }
                            value = property.GetValue(value, new object[] { descr.Index });
                        }                       
                    }
                }
                else
                {
                    value = property.GetValue(currentValue, new object[] { descr.Index });
                }

                currentValue = value;
                if ((descriptorLIst.Count - i) <= 1) // last item in the list == last property we need to return value
                {
                    break;
                }
            }

            return currentValue;
        }

        #endregion
        #endregion
    }
}
