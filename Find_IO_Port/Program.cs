using System.IO.Ports;
using System.Management;

namespace Find_IO_Port
{
    internal class Program
    {
        static void Main()
        {
            if (OperatingSystem.IsWindows())
            {
                //Options Menu
                Dictionary<int, string> options = new Dictionary<int, string>
                {
                    {1, "USB Detailed" },
                    {2, "Overview of USB devices and COM ports" },
                    {3, "To exit" }
                };

                //Format options
                string optionsString = string.Join(" or ", options.Select(option => $"'{option.Key}' for '{option.Value}'"));

                bool continueRunning = true;
                while (continueRunning)
                {
                    Console.WriteLine($"Enter {optionsString}");
                    string? opt = Console.ReadLine();

                    if (!string.IsNullOrEmpty(opt) && int.TryParse(opt, out int result))
                    {
                        Console.Write(Convert.ToInt32(opt) < 3 ? "=====================================\n" : string.Empty);

                        switch (Convert.ToInt32(opt))
                        {
                            case 1:
                                FindWindowsUSBDevices();
                                break;
                            case 2:
                                ListUSBDevices();
                                break;
                            case 3:
                                continueRunning = false;
                                break;
                            default:
                                Console.WriteLine("Invalid input.");
                                break;
                        }
                    }
                    else
                        Console.WriteLine("Invalid input.");
                }
            }
            else
            {
                Console.WriteLine("This feature is only supported on Windows.");
            }
        }

        #region option 1
        /// <summary>
        /// Detailed for USB-specific investigations, focusing on the USB controller relationship
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Windows-specific code")]
        static void FindWindowsUSBDevices()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_USBControllerDevice");

                if (searcher.Get().Count > 0)
                {
                    Console.WriteLine("\nDetailed USB Device specific investigations:\n");
                    Console.WriteLine("=====================================");
                }

                foreach (ManagementObject usbDevice in searcher.Get())
                {
                    string dependent = usbDevice["Dependent"].ToString()!;
                    string deviceID = dependent.Split(['='], 2)[1].Trim('"');

                    ManagementObjectSearcher pnpSearcher = new ManagementObjectSearcher(
                        $"SELECT * FROM Win32_PnPEntity WHERE DeviceID='{deviceID}'");

                    foreach (ManagementObject pnpDevice in pnpSearcher.Get())
                    {
                        string name = pnpDevice["Name"]?.ToString() ?? "Unknown Device";
                        string description = pnpDevice["Description"]?.ToString() ?? "No Description";

                        Console.WriteLine($"Device Name: {name}");
                        Console.WriteLine($"Description: {description}");
                        Console.WriteLine($"Device ID: {deviceID}");
                        Console.WriteLine();

                        if (pnpDevice["DeviceID"] != null)
                        {
                            string portInfo = pnpDevice["DeviceID"].ToString()!;
                            Console.WriteLine($"Port Information: {portInfo}");
                        }

                        Console.WriteLine("=====================================");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while retrieving USB device information: " + ex.Message);
            }
        }
        #endregion

        #region option 2
        /// <summary>
        /// Overview of USB devices, including COM port extraction, which is useful for serial communication tasks.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Windows-specific code")]
        static void ListUSBDevices()
        {
            try
            {
                // Get a list of serial port names
                string[] ports = SerialPort.GetPortNames();

                Console.Write(ports.Count() > 0 ? "Available COM Ports:\n" : string.Empty);
                foreach (string port in ports)
                {
                    Console.WriteLine("=====================================");
                    Console.WriteLine($"- {port}");
                    Console.WriteLine("=====================================");
                }

                Console.WriteLine("\nDetailed USB Device and COM port Information:\n");
                Console.WriteLine("=====================================");

                // Query Win32_PnPEntity for USB devices
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_PnPEntity WHERE PNPDeviceID LIKE '%USB%'");

                foreach (ManagementObject device in searcher.Get())
                {
                    string name = device["Name"]?.ToString() ?? "Unknown";
                    string deviceId = device["DeviceID"]?.ToString() ?? "No Device ID";
                    string description = device["Description"]?.ToString() ?? "No Description";

                    Console.WriteLine($"Device Name: {name}");
                    Console.WriteLine($"Description: {description}");
                    Console.WriteLine($"Device ID: {deviceId}");

                    // Extract COM port name if available
                    string comPort = ExtractCOMPortFromDevice(deviceId);

                    if (!string.IsNullOrEmpty(comPort))
                    {
                        Console.WriteLine($"COM Port: {comPort}");
                    }

                    Console.WriteLine("=====================================");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while retrieving USB device information: " + ex.Message);
            }
        }

        /// <summary>
        /// This function to extract COM port from device
        /// </summary>
        /// <param name="deviceId"></param>
        /// <returns></returns>
        static string ExtractCOMPortFromDevice(string deviceId)
        {
            // Using a regular expression to extract the COM port number
            var match = System.Text.RegularExpressions.Regex.Match(deviceId, @"\\(COM\d+)");

            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            return string.Empty;
        }
        #endregion
    }
}
