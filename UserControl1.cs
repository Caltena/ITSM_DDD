using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;

using System.Windows.Forms;
// Add references
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Win32;
using System.IO;

namespace ITSM_DDD
{

    [ProgId("ITSM_DropDMS")]
    [ClassInterface(ClassInterfaceType.AutoDual), ComSourceInterfaces(typeof(UserControlEvents))]
    public partial class DropDMS: UserControl
    {

        private string _strFiles;

        public string strFiles
        {
            get { return _strFiles; }
            set { _strFiles = value;  }
        }

        private string _strError;

        public string strError
        {
            get { return _strError; }
            set { _strError = value; }
        }



        public DropDMS()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form_DragEnter);
            this.DragDrop += new DragEventHandler(Form_DragDrop);
        }

        [ComVisible(false)]
        public void Form_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                string[] FileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);

                if (FileList.Count() > 0)
                {
                    string File = FileList[0];
                    if (CanFileAdd(File) == true)
                    {
                        strFiles = File;
                        DragDropEvent(1);
                    }
                    else
                    {
                        strError = string.Format("'{0}' ist eine Dateiordner.", File);
                        DragDropEvent(2);
                    }

                }
            } catch
            {
                strError = string.Format(" Error in Function 'Form_DragDrop'.");
                DragDropEvent(2);
            }
        }

        [ComVisible(false)]
        public Boolean CanFileAdd(string strFullPath)
        {
            FileAttributes attr = File.GetAttributes(strFullPath);
            if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
            {
                return false;
            }

            return true;
        }

        [ComVisible(false)]
        public void Form_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;                                    // Okay
            else
                e.Effect = DragDropEffects.None;                                // Unknown data, ignore it
        }



        #region register COM ActiveX object
        [ComRegisterFunction()]
        public static void RegisterClass(string key)
        {
            StringBuilder skey = new StringBuilder(key);
            skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            RegistryKey regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true);
            RegistryKey ctrl = regKey.CreateSubKey("Control");
            ctrl.Close();
            RegistryKey inprocServer32 = regKey.OpenSubKey("InprocServer32", true);
            inprocServer32.SetValue("CodeBase", Assembly.GetExecutingAssembly().CodeBase);
            inprocServer32.Close();
            regKey.Close();
        }


        [ComUnregisterFunction()]
        public static void UnregisterClass(string key)
        {
            StringBuilder skey = new StringBuilder(key);
            skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            RegistryKey regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true);
            regKey.DeleteSubKey("Control", false);
            RegistryKey inprocServer32 = regKey.OpenSubKey("InprocServer32", true);
            regKey.DeleteSubKey("CodeBase", false);
            regKey.Close();
        }
        #endregion

        #region  Eventhandler interface 
        public delegate void ControlEventHandler(int NumVal);
        [Guid("6A7C3A66-34AD-4513-BA6D-916D1F8B21B3")]
        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]


        public interface UserControlEvents
        {
            [DispId(0x60020001)]
            void DragDropEvent(int NumVal);
        }

        public event ControlEventHandler DragDropEvent;

   

        #endregion

        #region ActiveX Interface

        [ComVisible(true)]
        public interface ICOMCallable
        {
            string GetFileFolder();
            string GetError();
            void Clear();
            void WSize(int intWidth, int intHeight);
        }

        [ComVisible(true)]
        public string GetError()
        { 
            return strError;
        }


        [ComVisible(true)]
        public string GetFileFolder()
        {
            return strFiles;
        }

        [ComVisible(true)]
        public void Clear()
        {
            strFiles = "";
            strError = "";
        }

        [ComVisible(true)]
        public void WSize(int intWidth ,int  intHeight)
        {
            try
            {
                this.Height = intHeight;
                this.Width = intWidth;
                this.Size = new System.Drawing.Size(intWidth, intHeight);
            }
            catch (Exception e)
            {
                strError = e.Message.ToString();
            }

        }
        #endregion

    }
}
