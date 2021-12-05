using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using USBPlugEvent.Properties;

namespace USBPlugEvent
{
    internal class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        [STAThread]
        static void Main(string[] args)
        {
            var arguments = args.ToList();

            if (!arguments.Any() || arguments.Contains("-h"))
            {
                Help();
                return;
            }

            if (arguments.Contains("-l"))
            {
                ListUSBDevices();
                return;
            }

            if (arguments.Contains("-d"))
            {
                ListUSBDeviceListener();
                return;
            }

            var onInsert = arguments.Contains("-i");
            var onRemove = arguments.Contains("-r");
            if (onInsert == false && onRemove == false)
            {
                onInsert = true;
                onRemove = true;
            }
            
            var parameters = arguments.Where(p => !p.StartsWith("-"));
            
            if (parameters.Count() == 2)
            {
                var device_id = (string)parameters.ElementAt(0);
                var cmd = (string)parameters.ElementAt(1);

                var handle = GetConsoleWindow();
                ShowWindow(handle, SW_HIDE);

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new USBListenerApplicationContext(onInsert, onRemove, device_id, cmd));
                
            }
            return;
        }

        private static void Help()
        {
            Console.WriteLine($"");
            Console.WriteLine($"Usage: USBPlugEvent.exe [-l] [-d] [-i] [-r] \"device_id\" \"cmd_to_launch\"");
            Console.WriteLine($"");
            Console.WriteLine($"Example: USBPlugEvent.exe -r \"USB\\VID_056D&PID_C07E\\4687336F3936\" \"msg %username% the specified device was removed\"");
            Console.WriteLine($"");
            Console.WriteLine($"Options:");
            Console.WriteLine($"    -l        List all currently connected USB devices.");
            Console.WriteLine($"    -d        Listen to Inserted/Removed USB devices.");
            Console.WriteLine($"    -i        Launch the command on insert only.");
            Console.WriteLine($"    -r        Launch the command on remove only.");
        }

        private static void ListUSBDeviceListener()
        {
            Console.WriteLine($"");
            Console.WriteLine($"DeviceID | Description | Event");

            WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
            ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
            insertWatcher.EventArrived += new EventArrivedEventHandler(ListUSBDeviceListenerDeviceInsertedEvent);
            insertWatcher.Start();

            WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
            ManagementEventWatcher removeWatcher = new ManagementEventWatcher(removeQuery);
            removeWatcher.EventArrived += new EventArrivedEventHandler(ListUSBDeviceListenerDeviceRemovedEvent);
            removeWatcher.Start();

            // Do something while waiting for events
            System.Threading.Thread.Sleep(2000000000);
        }

        private static void ListUSBDeviceListenerDeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            Console.WriteLine("{0} | {1} | Inserted", instance.GetPropertyValue("DeviceID"), instance.GetPropertyValue("Description"));
        }

        private static void ListUSBDeviceListenerDeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            Console.WriteLine("{0} | {1} | Removed", instance.GetPropertyValue("DeviceID"), instance.GetPropertyValue("Description"));
        }

        static void ListUSBDevices()
        {
            ManagementObjectCollection collection;
            using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_USBHub"))
            collection = searcher.Get();
            Console.WriteLine($"");
            Console.WriteLine($"DeviceID | Description");
            foreach (var device in collection)
            {
                //Console.WriteLine("Device ID: {0}, PNP Device ID: {1}, Description: {2}", device.GetPropertyValue("DeviceID"), device.GetPropertyValue("PNPDeviceID"), device.GetPropertyValue("Description"));
                Console.WriteLine("{0} | {1}", device.GetPropertyValue("DeviceID"), device.GetPropertyValue("Description"));
            }
            collection.Dispose();
        }
    }

    public class USBListenerApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;

        private static bool _onInsert = true;
        private static bool _onRemove = true;
        private static string _device_id = string.Empty;
        private static string _cmd = string.Empty;

        public USBListenerApplicationContext(bool onInsert, bool onRemove, string device_id, string cmd)
        {
            _onInsert = onInsert;
            _onRemove = onRemove;
            _device_id = device_id;
            _cmd = cmd;

            // Initialize Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Resources.AppIcon,
                ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };
            
            USBSpecificDeviceListener();
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;

            Application.Exit();
        }

        private static void USBSpecificDeviceListener()
        {
            if (_onInsert)
            {
                WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
                ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
                insertWatcher.EventArrived += new EventArrivedEventHandler(SpecificDeviceInsertedEvent);
                insertWatcher.Start();
            }

            if (_onRemove)
            {
                WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBHub'");
                ManagementEventWatcher removeWatcher = new ManagementEventWatcher(removeQuery);
                removeWatcher.EventArrived += new EventArrivedEventHandler(SpecificDeviceRemovedEvent);
                removeWatcher.Start();
            }
        }

        private static void SpecificDeviceInsertedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            if (instance.GetPropertyValue("DeviceID").ToString() == _device_id)
            {
                LaunchCmd(_cmd);
            }
        }

        private static void SpecificDeviceRemovedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            if (instance.GetPropertyValue("DeviceID").ToString() == _device_id)
            {
                LaunchCmd(_cmd);
            }
        }

        private static void LaunchCmd(string cmd)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C " + cmd;
            process.StartInfo = startInfo;
            process.Start();
        }
    }
}
