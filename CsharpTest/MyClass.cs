using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.Advertisement;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Enumeration;
using System.Runtime.InteropServices;
using System.IO;

namespace CsharpTest
{
    [ComVisible(true)]
    public interface IMyClass
    {
        string Main();
    }
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class MyClass : IMyClass
    {
        private const string SignalStrengthProperty = "System.Devices.Aep.SignalStrength";
        static SortedDictionary<string, int> myDict = new SortedDictionary<string, int>();
        static int flag;
        static string blueToothInfor;
        static string blueToothDevice1 = "WPS1111";
        static string blueToothDevice2 = "WPS2222";
        //public void Main()
        //{
        //    flag = 0;
        //    StartScanning();
        //}

        public string Main()
        {
            //try
            //{
                // 測試Error log有沒有寫到.txt檔
                //Exception e = new Exception("write error log success!!!!!!!! wow~~ 66666666");
                //ErrorLogging(e);


                flag = 0;
                blueToothInfor = "";
           
                // Query for extra properties you want returned
                string[] requestedProperties = { "System.Devices.Aep.DeviceAddress", "System.Devices.Aep.IsConnected" };
                var BluetoothDeviceSelector = "System.Devices.DevObjectType:=5 AND System.Devices.Aep.ProtocolId:=\"{E0CBF06C-CD8B-4647-BB8A-263B43F0F974}\"";
                var additionalProperties = new[] { SignalStrengthProperty };
                DeviceWatcher deviceWatcher = DeviceInformation.CreateWatcher(BluetoothDeviceSelector, additionalProperties);

            
                // Register event handlers before starting the watcher.
                // Added, Updated and Removed are required to get all nearby devices
                deviceWatcher.Added += DeviceWatcher_Added;
                

                deviceWatcher.Updated += DeviceWatcher_Updated;
                //deviceWatcher.Removed += DeviceWatcher_Removed;
           

                // EnumerationCompleted and Stopped are optional to implement.
                deviceWatcher.EnumerationCompleted += DeviceWatcher_EnumerationCompleted;
                deviceWatcher.Stopped += DeviceWatcher_Stopped;
                Thread.Sleep(1000);

                //Start the watcher.
                deviceWatcher.Start();
                flag = 0;
                blueToothInfor = "";
                while (flag == 0)
                {

                }
            //}
            //catch (Exception ex)
            //{
                //ErrorLogging(ex);
            //}


            return blueToothInfor;
        }

        public void DeviceWatcher_Stopped(DeviceWatcher sender, object args)
        {
            foreach (var item in myDict)
            {
                if (item.Key == blueToothDevice1 || item.Key == blueToothDevice2)
                {
                    //Console.WriteLine(item.Key + " " + item.Value);
                    blueToothInfor = blueToothInfor + item.Value.ToString() + " ";
                }
            }
            myDict.Clear();
            flag = 1;
            //sender.Start();
        }

        private static void DeviceWatcher_EnumerationCompleted(DeviceWatcher sender, object args)
        {
            sender.Stop();
            //throw new NotImplementedException();
        }

        //private static void DeviceWatcher_Removed(DeviceWatcher sender, DeviceInformationUpdate args)
        //{
        //    //throw new NotImplementedException();
        //}

        private static void DeviceWatcher_Updated(DeviceWatcher sender, DeviceInformationUpdate args)
        {
            if (myDict.ContainsKey(blueToothDevice1) && myDict.ContainsKey(blueToothDevice2))
            {
                sender.Stop();
            }
            //throw new NotImplementedException();
        }

        private static void DeviceWatcher_Added(DeviceWatcher sender, DeviceInformation args)
        {

            //try
            //{
                //Exception e1 = new Exception("go DeviceWatcher_Added");
                //ErrorLogging(e1);

                myDict.Add(args.Name, Convert.ToInt16(args.Properties[SignalStrengthProperty]));

                //Exception e2 = new Exception("leave DeviceWatcher_Added");
                //ErrorLogging(e2);
            //}
            //catch (Exception ex)
            //{
                //ErrorLogging(ex);
            //}

           
        }

        //export error log
        public static void ErrorLogging(Exception ex)
        {
            string strPath = @"C:\aaa\Log.txt";  //忘記相對路徑怎麼寫了，先寫絕對路徑
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);

            }
        }
    }
}