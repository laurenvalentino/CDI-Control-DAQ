"""
Copyright © 2026, UChicago Argonne, LLC
All Rights Reserved
 Software Name: Supervisory Control and Data Acquisition for Electrochemical Separation Experimentation
By: Argonne National Laboratory
BSD OPEN SOURCE LICENSE

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.


******************************************************************************************************
DISCLAIMER

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
***************************************************************************************************

"""

"""
Author: Eric Vu
Mentor: Lauren Valentino

Purpose: 
    This program defines a Python class, `SerialDevice`, for controlling a probe device 
    used in Capacitive Deionization (CDI) experiments via serial communication. The class 
    manages the connection to the probe through a specified COM port and includes methods 
    to send various commands and retrieve data.
    
Methods: 
    Key functionalities of the `SerialDevice` class include:
        - Establishing a serial connection with the probe using parameters like baud rate, byte size, stop bits, and parity.
        - Sending commands to the probe and receiving responses, with methods for specific operations such as:
        - `send_help()`: Displays the help menu of the probe device.
        Conductivity Commands: 
            - `send_zero()`: Turns off zeroing of the probe.
            - `send_range(value)`: Sets the range of the probe, with auto-ranging based on the provided value.
            - `send_calc(value)`: Sends a calculation command to the probe with a specific value (e.g., kappa value for conductivity).
            - `send_show()`: Displays the currently set kappa value.
            - `send_v()`: Retrieves and processes the probe's voltage reading.
            - `send_binary(value)`: Samples data from the probe in binary format.
            - Extracting numerical values from the probe's output using `extract_numeric_value()`.
            - Safely closing the serial connection with the `close()` method.
            
"""

import serial
import time

class SerialDevice:
    BAUDRATE = 115200
    TIMEOUT = 1
    BYTE_SIZE = serial.EIGHTBITS
    STOP_BITS = serial.STOPBITS_ONE
    PARITY = serial.PARITY_NONE
    RESPONSE_DELAY = 0.5
    
    def __init__(self, port):
        self.port = port
        self.ser = None

    def connect(self):
        if self.ser is None or not self.ser.is_open:
            try:
                self.ser = serial.Serial(
                    port=self.port,
                    baudrate=self.BAUDRATE, 
                    timeout=self.TIMEOUT,
                    bytesize=self.BYTE_SIZE,
                    stopbits=self.STOP_BITS,
                    parity=self.PARITY
                )
            except serial.SerialException as e:
                raise RuntimeError(f"Failed to connect to {self.port}: {e}")

    def send_command(self, command):
        try:
            self.connect()
            self.ser.write((command + '\r\n').encode('utf-8'))
            time.sleep(self.RESPONSE_DELAY)
            response = self.ser.read_all().decode('utf-8')
            return response.strip()
        except serial.SerialException as e:
            return f"Error: {e}"
        
    def extract_numeric_value(self, probe_output):
        # Extract the portion of the string between the first character and 'EPU'
        extracted_value = probe_output[1:probe_output.index('EPU')]
        return float(extracted_value)

    def close(self):
        if self.ser is not None and self.ser.is_open:
            self.ser.close()

    def send_help(self):
        return self.send_command('help')
    
    def send_zero(self):
        return self.send_command('zero off')
    
    def send_range(self, value):
        return self.send_command(f'set range {value} auto')

    def send_calc(self, value):
        return self.send_command(f'calc k {value}')

    def send_show(self):
        return self.send_command('show k')
    
    def send_v(self):
        probe_output = self.send_command('v')
        return self.extract_numeric_value(probe_output)

    def send_binary(self, value):
        return self.send_command(f'sample binary {value}')

    def send_set(self, value):
        return self.send_command('set k {value}')


if __name__ == "__main__":
    port = 'COM3'
    probe = SerialDevice(port)

    try:
        print(probe.send_help())
        print(probe.send_zero())
        print(probe.send_range(2))
        print(probe.send_calc(1.413))
        print(probe.send_show())
        print(probe.send_v())
        print(probe.send_binary(1))

        result = probe.send_calc(1.413)
        result2 = probe.send_show()
        result3 = probe.send_v()
 
    finally:
        probe.close()



#zeroing - potentially
#set range - confirm samantha when open podvu select 200,20,2,0.2 top RHS value confirm value / possibly select auto 
#calc k - stored internally, record k value in spreadsheet, kappa: 1.413 miliSiemens (given)
#show k - record in spreadsheet 
    #if needed, set k - only when pc restarts, or usb connection is resetted, 
#v 
#sample - read data 1 second interval
