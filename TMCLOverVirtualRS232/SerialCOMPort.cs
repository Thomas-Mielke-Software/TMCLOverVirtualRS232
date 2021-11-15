/*----------------------------------------------------------------------------------------------------------------------
 * C# .NET Class to generate TMCL commands which can be sent to a TRINAMIC steprocker hardware device
 * Created by Wolfgang Kurz { Domain: http://wkurz.com,  E-Mail wolfgang@uwkurz.de }
 * Copyright: (c) Wolfgang Kurz, 2017
 * 
 * This class can be used 'as is' for free. It can also be modified by anybody who has a need to do so.
 * The author and TRINAMIC Motion Control is not responsible for any problems caused by potential errors in this class.
 *----------------------------------------------------------------------------------------------------------------------
 */

using System;
using System.IO.Ports;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TMCLOverVirtualRS232
{
    public partial class SerialCOMPort
    {
        private string[] _availablePorts = System.IO.Ports.SerialPort.GetPortNames();
        private int _noOfPorts;
        private string _selectedPortName = "";
        public static SerialPort activePort = null;
        private static bool _isOpen = false;
        public static SerialCOMPort activeInstance = null;

        /// <summary>
        /// Creates a new instance of this class
        /// </summary>
        public SerialCOMPort()
        {
            activeInstance = this;         // for easy reference to this instance
            _availablePorts = AvailablePortNames;
            _noOfPorts = _availablePorts.Length;
        }

        /// <summary>
        /// Evaluates the available Serial Ports at a computer
        /// </summary>
        public string[] AvailablePortNames
        {
            get { return _availablePorts; }
        }

        /// <summary>
        /// Returns the name of the selected Port as string
        /// </summary>
        public string SelectedPortName { get; set; }


        /// <summary>
        /// Gets the connection status of a Serial Port identified via its port name
        /// </summary>
        /// <param name="aPortName"></param>
        /// <returns>true as connected and false as not connected</returns>
        public bool GetConnection(String aPortName)
        {
            bool retVal = false;
            int i;
            for (i = 0; i < _noOfPorts; i++)
            {
                if (_availablePorts[i].Contains(aPortName.Trim()) == true)
                {
                    retVal = true;
                }
            }
            return retVal;
        }

        /// <summary>
        /// Gets the open status of the active serial port
        /// </summary>
        /// <returns>'true' as OPEN or 'false' as CLOSED</returns>
        public bool GetOpenState()
        {
            return _isOpen;
        }

        /// <summary>
        /// Sets the selected Port with the index of the list of the available Ports
        /// </summary>
        /// <param name="aIndex"></param>
        /// <returns>'true' = port selected,  'false' = port selection failed</returns>
        public bool SetSelectedPort(int aIndex)
        {
            bool returnValue = false;
            string newPortName;

            try
            {
                newPortName = _availablePorts[aIndex];

                if (_selectedPortName != newPortName)
                {
                    if (activePort != null)  // Port is open !
                    {
                        CloseActivePort();
                    }

                    _selectedPortName = newPortName;
                    activePort = new SerialPort(_selectedPortName);
                    returnValue = true;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception)
            {
                returnValue = false;
            }

            return returnValue;
        }

        /// <summary>
        /// Opens the active Port
        /// </summary>
        public void OpenPort()
        {
            if (activePort == null)
            {
                activePort = new SerialPort(_selectedPortName);
            }
            else
            {
                // do nothing port exists
            }
            activePort.PortName = _selectedPortName;
            activePort.BaudRate = 9600;
            activePort.ReadTimeout = 5000;
            activePort.WriteTimeout = 5000;
            activePort.WriteTimeout = 300;

            // set the appropriate properties of the com port
            activePort.ReadBufferSize = 40;
            activePort.WriteBufferSize = 20;

            //try   -- let the outside world handle exceptions
            //{
            if (_isOpen == false)
            {
                activePort.Open();
                _isOpen = true;
            }
            //}
            //catch (UnauthorizedAccessException uae)
            //{
            //    Console.WriteLine("SerialCOMPort 0232  UnauthorizedAccesssException =" + uae);
            //    MessageBox.Show(Language.GetM(119), Language.GetM(99), MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            //}
        }

        /// <summary>
        /// send an array of bytes as TMCM command to the active COM Port
        /// </summary>
        /// <returns></returns>
        public int SendTMCMCommand(int cmdId, int anAddress, int aType, int aBank, int aValue)
        {
            int result = -1;
            byte[] readBuffer = new byte[9];
            byte[] commandBytes = new byte[9];

            if (activePort != null)
            {
                if (_isOpen == false)
                {
                    try
                    {
                        activePort.Open();
                        _isOpen = true;
                    }
                    catch (Exception)
                    {
                        _isOpen = false;
                    }
                }
                if (_isOpen == true)
                {
                    commandBytes = setTMCMCommand(cmdId, anAddress, aType, aBank, aValue);
                    // presentBytes(commandBytes);      // Test statement

                    activePort.Write(commandBytes, 0, 9);
                    Thread.Sleep(1000);
                    activePort.Read(readBuffer, 0, 9);
                    result = Convert.ToInt32(readBuffer[2]);

                    Console.WriteLine($"SendTMCMCommand execution successful -- result status code {result}: {interpretStatusCode(result)}");
                }
                    
            }
            return result;
        }

        private void presentBytes(byte[] commandBytes)
        {
            Console.WriteLine("SerialCOMPort 0306 cammandBytes = [0] " + commandBytes[0] + ", [1] "
                                                                       + commandBytes[1] + ", [2] "
                                                                       + commandBytes[2] + ", [3] "
                                                                       + commandBytes[3] + ", [4] "
                                                                       + commandBytes[4] + ", [5] "
                                                                       + commandBytes[5] + ", [6] "
                                                                       + commandBytes[6] + ", [7] "
                                                                       + commandBytes[7] + ", [8] "
                                                                       + commandBytes[8] + " = checksum ;");
        }

        /// <summary>
        /// Close the active Port and set the indicator to false which is 'closed'
        /// </summary>
        /// <returns></returns>
        public void CloseActivePort()
        {
            activePort.Close();
            _isOpen = false;
        }

        /// <summary>
        /// Sets the TMCM command based on the instruction Id using the type and a value
        /// </summary>
        /// <param name="cmdId, anAddress, aType, intValue"></param>
        /// <returns></returns>
        private byte[] setTMCMCommand(int cmdId, int anAddress, int aType, int aBank, int intValue)
        {
            byte aCmdIdByte = (byte)cmdId;
            byte[] theTMCMCommand = new byte[9];
            byte[] valueBytes = new byte[4];
            byte addressByte = (byte)anAddress;
            byte typeByte = (byte)aType;
            byte aBankByte = (byte)aBank;
            int corVal = 0;

            if (intValue < 0) { corVal = -1; } else { corVal = 0; }
            corVal = corVal + 0;

            valueBytes = convertIntToBytes(intValue); // Converts the provided integer value into a 4 byte array and computes it's checksum

            theTMCMCommand = setTheBytes(addressByte, aCmdIdByte, typeByte, aBankByte, valueBytes);
            theTMCMCommand[8] = setTheCheckSum(theTMCMCommand);

            return theTMCMCommand;
        }

        /// <summary>
        /// set the command  bytes
        /// </summary>
        /// <returns></returns>
        private byte[] setTheBytes(byte addressByte, byte aCmdIdByte, byte typeByte, byte aBankByte, byte[] valueBytes)
        {
            byte[] theTMCMCommand = new byte[9];

            theTMCMCommand[0] = addressByte;      // the target address
            theTMCMCommand[1] = aCmdIdByte;             // the instruction number
            theTMCMCommand[2] = typeByte;         // type
            theTMCMCommand[3] = aBankByte;        // motor bank 
            theTMCMCommand[4] = valueBytes[0];
            theTMCMCommand[5] = valueBytes[1];
            theTMCMCommand[6] = valueBytes[2];
            theTMCMCommand[7] = valueBytes[3];
            theTMCMCommand[8] = 0x00;

            return theTMCMCommand;
        }

        /// <summary>
        /// set the command  bytes
        /// </summary>
        /// <returns></returns>
        private byte setTheCheckSum(byte[] theTMCMCommand)
        {
            uint checksum;
            byte checksumByte;
            uint checksum24;
            checksum = (uint)theTMCMCommand[0] + (uint)theTMCMCommand[1] + (uint)theTMCMCommand[2] + (uint)theTMCMCommand[3] +
                    (uint)theTMCMCommand[4] + (uint)theTMCMCommand[5] + (uint)theTMCMCommand[6] + (uint)theTMCMCommand[7];
            checksum24 = Convert.ToUInt32(checksum << 24);
            checksumByte = Convert.ToByte(checksum24 >> 24);
            return checksumByte;
        }

        /// <summary>
        /// convert an integer into a  4 byte long byte array 
        /// </summary>
        /// <returns></returns>
        private byte[] convertIntToBytes(int aInt)
        {
            int[] value = new int[4] { 0, 0, 0, 0 };    // integer with 32 valid bits
            char[] aChar = new char[4] { '0', '0', '0', '0' };   // integer with 32 valid bits
            int[] val8b = new int[4] { 0, 0, 0, 0 };    // integer with 8 rightmost valid bits = 0 1o 255
            uint theInt = (uint)aInt;
            uint valTmp = theInt;
            uint aTemp;

            aTemp = (uint)aInt;

            aChar[0] = Convert.ToChar(theInt >> 24);
            valTmp = theInt << 8;
            aChar[1] = Convert.ToChar(valTmp >> 24);
            valTmp = theInt << 16;
            aChar[2] = Convert.ToChar(valTmp >> 24);
            valTmp = theInt << 24;
            aTemp = (uint)valTmp >> 24;
            aChar[3] = Convert.ToChar(aTemp >> 24);

            byte[] byteArray = new byte[4] { Convert.ToByte(aChar[0]), Convert.ToByte(aChar[1]), Convert.ToByte(aChar[2]), Convert.ToByte(aChar[3]) };

            // Console.WriteLine("0396 SerialCOMPort byteArray = " + byteArray[0] + "," + byteArray[1] + "," + byteArray[2] + "," + byteArray[3] ); // Test stetemnet

            return byteArray;
        }

        private string interpretStatusCode(int status)
        {
		    switch (status)
            {
                case 100 : return "Success";
                case 101: return "Command loaded";
                case 1: return "Incorrect Checksum";
                case 2: return "Invalid Command";
                case 3: return "Wrong Type";
                case 4: return "Invalid Value";
                case 5: return "EEPROM Locked";
                case 6: return "Command not Available";
            }
            return "Unknown Status Code";
        }
    }
}