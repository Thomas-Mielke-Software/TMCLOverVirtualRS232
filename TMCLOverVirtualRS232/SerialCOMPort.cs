/*----------------------------------------------------------------------------------------------------------------------
 * C# .NET Class to generate TMCL commands which can be sent to a TRINAMIC
 * steprocker hardware device
 *
 * Copyright (c) Wolfgang Kurz, 2017
 * Copyright (c) Thomas Mielke, 2021
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published
 * by the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNULesser General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
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
        /// <param name="cmdId">TMCL command like 9 for 'Set Global Parameter'</param>
        /// <param name="anAddress">Address of the TMCM device</param>
        /// <param name="aType">Type as submitted in byte 2 of the command buffer</param>
        /// <param name="aBank">Motor bank</param>
        /// <param name="aValue">32-bit value submitted in byte 4-7</param>
        /// <returns>status code as from byte 2 of the response buffer</returns>
        public int SendTMCMCommand(int cmdId, int anAddress, int aType, int aBank, int aValue)
        {
            int result = -1000;
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
                    Thread.Sleep(1);
                    activePort.Read(readBuffer, 0, 9);
                    result = Convert.ToInt32(readBuffer[2]);

                    Console.WriteLine($"SendTMCMCommand execution successful -- result status code {result}: {InterpretStatusCode(result)}");
                }
            }
            return result;
        }

        /// <summary>
        /// send an array of bytes as TMCM command to the active COM Port
        /// </summary>
        /// <param name="cmdId">TMCL command like 9 for 'Set Global Parameter'</param>
        /// <param name="anAddress">Address of the TMCM device</param>
        /// <param name="aType">Type as submitted in byte 2 of the command buffer</param>
        /// <param name="aBank">Motor bank</param>
        /// <param name="aValue">32-bit value submitted in byte 4-7</param>
        /// <param name="aReturnedValue">32-bit value returned, for example if using cmdID 10 'Get Global Parameter'</param>
        /// <returns>status code as from byte 2 of the response buffer</returns>
        public int SendTMCMCommand(int cmdId, int anAddress, int aType, int aBank, int aValue, out int aReturnedValue)
        {
            int result = -1000;
            aReturnedValue = 0;
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
                    Thread.Sleep(1);
                    activePort.Read(readBuffer, 0, 9);
                    result = Convert.ToInt32(readBuffer[2]);
                    var intValueInBytes = new byte[4];
                    intValueInBytes[0] = readBuffer[4];
                    intValueInBytes[1] = readBuffer[5];
                    intValueInBytes[2] = readBuffer[6];
                    intValueInBytes[3] = readBuffer[7];
                    aReturnedValue = convertBytesToInt(intValueInBytes);

                    Console.WriteLine($"SendTMCMCommand execution successful -- result status code {result}: {InterpretStatusCode(result)}");
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
            theTMCMCommand[1] = aCmdIdByte;       // the instruction number
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
        /// convert an integer into a 4 byte long byte array 
        /// </summary>
        /// <returns></returns>
        private byte[] convertIntToBytes(int aInt)
        {
            byte[] intBytes = BitConverter.GetBytes(aInt);
            if (BitConverter.IsLittleEndian)
                Array.Reverse(intBytes);        // convert to little endian byte order
            return intBytes;
        }

        /// <summary>
        /// convert 4 byte long byte array to integer 
        /// </summary>
        /// <returns></returns>
        private int convertBytesToInt(byte[] aBuffer)
        {
            if (aBuffer.Length != 4) return 0;  // invalid buffersize
            if (BitConverter.IsLittleEndian)
                Array.Reverse(aBuffer);         // convert to little endian byte order
            int i = BitConverter.ToInt32(aBuffer, 0);
            return i;
        }

        /// <summary>
        /// shows a string representation of the given status code
        /// </summary>
        /// <returns>human readable text</returns>
        public string InterpretStatusCode(int status)
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