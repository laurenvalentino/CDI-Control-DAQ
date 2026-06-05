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
Author: Molly Perez
Mentor: Lauren Valentino 

Purpose: 
    This program defines a Python class, 'psu', for controlling a power supply
    used in Capacitive Deionization (CDI) experiments via serial communication. The class 
    manages the connection to the pump through USB includes methods to send commands and retrieve data.
    
Methods:
    Key functionalities of the `PSU` class include: 
        Establishing connection via USB by a string.
        Sending commands to the probe and receiving responses, with methods for specific operations such as:
            
        Power Supply Commands:
            -'set_voltage_curren()': Applys the voltage and current for the power supply with the given user input, hardcoded for Channel 1
            -'measure_voltage())':  Measure the actual voltage of the power supply's channel
            -'measure_current()':   Measure the actual current of the power supply's channel
            -'set_voltage_current_prg4()': 
            -' measure_voltage_prg4()':  Measure the actual voltage of the power supply's for another channel 
            -'measure_current_prg4()':  Measure the actual Current of the power supply's for another channel 
            -'close_Program4()':  Closes the connection properly turning off all the channels 
            -'close()' : Closes the connection properly turning off the channel
"""

import pyvisa


class psu:
    
    def __init__(self, resource):
        self.rm = pyvisa.ResourceManager()
        self.instr = self.rm.open_resource(resource)
    
    def set_voltage_current(self, voltage, current_mA): 
        
        # Convert voltage and current_mA to floats if they are strings 
        current_mA = float(current_mA) 
        # Convert milliamps to amps 
        current_A = current_mA / 1000.0 
        # Turn on the output channel 
        self.instr.write(':OUTP CH1,ON') 
        # Create the command string and write it to the instrument 
        command = f':APPL CH1,{voltage},{current_A:.2f}' 
        self.instr.write(command) 
        # Print the result 
        print(f'Set channel 1 to {voltage}V and {current_A:.2f}A') 
        
    def measure_voltage(self):
        query_command = ':MEAS? CH1'
        actual_voltage = self.instr.query(query_command).strip()  # Remove newline or whitespace
        return actual_voltage
    
    def measure_current(self):
        query_command = ':MEAS:CURR? CH1'
        actual_current = self.instr.query(query_command).strip()  # Remove newline or whitespace
        actual_current = float(actual_current)  # Convert to float
        actual_current = actual_current * 1000  # Convert to mA
        return actual_current
    
    def close(self):
        self.instr.write(':OUTP CH1,OFF')
        self.instr.write(':APPL CH1,0,0')
        self.instr.close()
        self.rm.close()
        
    def set_voltage_current_prg4(self, channel, voltage, current_mA):
        current_mA = float(current_mA)
        current_A = current_mA / 1000.0
        self.instr.write(f':OUTP {channel},ON')
        command = f':APPL {channel},{voltage},{current_A:.2f}'
        self.instr.write(command)
        print(f'Set {channel} to {voltage}V and {current_A:.2f}A')

    def measure_voltage_prg4(self, channel):
        query_command = f':MEAS? {channel}'
        actual_voltage = self.instr.query(query_command).strip()
        return actual_voltage

    def measure_current_prg4(self, channel):
        query_command = f':MEAS:CURR? {channel}'
        actual_current = self.instr.query(query_command).strip()
        actual_current = float(actual_current) * 1000  # Convert to mA
        return actual_current

    def close_Program4(self):
        self.instr.write(':OUTP CH1,OFF')
        self.instr.write(':OUTP CH2,OFF')
        self.instr.write(':APPL CH1,0,0')
        self.instr.write(':APPL CH2,0,0')
        self.instr.close()
        self.rm.close()

def main():
    resource = 'USB0::6833::3601::DP8C203202681::0::INSTR'  # Replace with your instrument's resource string
    dp832 = psu(resource)

    # Here you can perform voltage and current measurements, settings, etc.
    voltage = dp832.measure_voltage()
    current = dp832.measure_current()
    print(f"Measured voltage: {voltage}V")
    print(f"Measured current: {current}A")
    
    dp832.close()

if __name__ == '__main__':
    main()
