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
    This program defines a Python class, `SerialDevice`, for controlling a Multi Function isoPode device 
    used in Capacitive Deionization (CDI) experiments via serial communication. The class 
    manages the connection to the probe through a specified COM port and includes methods 
    to send various commands and retrieve data.

    Website: https://www.edaq.com/product_sheets/software/ES350_PodVu_Software.pdf
    
Methods: 
    Key functionalities of the `SerialDevice` class include:
        - Establishing a serial connection with the probe using parameters like baud rate, byte size, stop bits, and parity.
        - Sending command() sends commands to the probe and receive responses, with methods for specific operations such as:
            pH Commands: 
                - set_ph_channel(): Sets the pH probe to a specific channel 
                - 'set_range_ph()': Setting a specific range for readings
                - 'get_range_ph()': Get the range to verify it's correctly set
                - 'set_unit_ph()':  Set unti to ph 
                - 'set_calibration_ph()': Setting calibration for specific solution given ph and mV values.
                - 'calibration_info_ph()': returning the calibration information as well returns the calubration slope and y-intercept   
            Conductivity Commands: 
                - 'set_conductivity_channel()': Sets the conductivity probe to a specific channel
                - 'send_zero_conductivity ()':  Removes any input offset
                - 'send_range_conductivity()':  Set measurement range idealy its going to be range 2 1 - 200 µS/cm
                - 'send_set_conductivity()':    Sets the k value of the conductivity probe with the users specific value
                - 'get_k_value()':              Print K value only no other string
                - 'send_calc_conductivity()':   Take a reading and calculate the ocrrect K to give the specified conductivity value (m)
            Temperature Commands:
                - 'set_temp_channel()': Setting temp probe to the appropriate channel (channel 3)
                - 'set_range_temp()':   Setting reading range for ph probe
                - 'get_range_temp()':   Getting the temp range that is applied to probe
                - 'get_ranges_temp()':  Optinal, printing ranges available
                - 'get_reading_temp()': Get a reading on the temp probe
            Others: 
                - 'set_off_channel()': Setting channel 4 as off, not in use. 
                -'send_readings()':sends reading of pH, Conductivity and temperature
                - 'get_raw_probe_values()': Gets raw probe values of the probe readings
                - 'get_ph_calibration_values()': Only retrives the solpe and y intercept of the calibration info
                - 'clear_all_inputs()': Clears all inputs of the pH calibration
                - 'close()': Closes the connection of the multi edaq podvu
"""


import serial
import time
import re


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
            response = self.ser.read_all().decode('latin-1')  # <-- FIXED
            return response.strip()
        except serial.SerialException as e:
            return f"Error: {e}"

    def close(self):
        if self.ser is not None and self.ser.is_open:
            self.ser.close()          
#--------------------------Ph & ISE commands ---------------------------------#           
    
    def set_ph_channel(self):
        return self.send_command('set channel 1 function pH&ISE')
        
    def set_range_ph(self,value):
        return self.send_command(f'set channel 1 range {value}')
    
    def get_range_ph(self):
        return self.send_command('get channel 1 range')
        
    def set_unit_ph(self):
        return self.send_command('set channel 1 units ')
    
    def set_calibration_ph (self, point_num, ph_value, mv_value):
        return self.send_command(f' cal channel 1 set {point_num} {ph_value} {mv_value}')
    
    def calibration_info_ph (self):
        return self.send_command('cal channel 1 get')
            
 #--------------------------Ph & ISE commands ---------------------------------#           
    
#--------------------------Conductivity commands -----------------------------#
    def set_conductivity_channel(self):
        return self.send_command('set channel 2 function Cond')
    
    def send_zero_conductivity (self):
        return self.send_command('set channel 2 offset off')
    
    def send_range_conductivity(self, value):
        return self.send_command(f'set channel 2 range {value} ')
    
    def send_set_conductivity (self, value): 
        return self.send_command(f'set channel 2 k {value}')
    
    def get_k_value(self):
        raw_response = self.send_command('get channel 2 K')
        print(f"Raw K response: {raw_response}")
    
        match = re.search(r'k\s+([\d.]+)', raw_response)
        if match:
            try:
                k_value = float(match.group(1))
                return k_value
            except ValueError:
                print("Error: K value found but could not convert to float.")
                return None
        else:
            print("Error: Could not find K value in response.")
            return None
    
    def send_calc_conductivity(self, value):
        return self.send_command(f'calc channel 2 k {value}')
    
#--------------------------Conductivity commands -----------------------------#       

#----------------------------- Pt 1K RTD commands -----------------------------# 

    def set_temp_channel (self):
        return self.send_command('set channel 3 function RTD')
    
    def set_range_temp(self,value):
        return self.send_command(f'set channel 3 range {value}')
    
    def get_range_temp(self):
        return self.send_command('get channel 3 range')
    
    def get_ranges_temp(self):
        return self.send_command('get channel 3 ranges')
    
    def get_reading_temp (self):
        return self.send_command('get channel 3 ohms')

#----------------------------- Pt 1K RTD commands -----------------------------#  
    def set_off_channel (self):
        return self.send_command('set channel 4 function Off')
       
#--------------------------- Retriving_data -----------------------------------#

    def send_readings(self, yintercept, slope, k_value):
        probe_output = self.send_command('adc')
        print(f"Raw probe output: {probe_output}")
    
        try:
           
            voltage_match = re.search(r"([-+]?\d*\.?\d+)\s*V", probe_output)
            cond_match = re.search(r"([\d.]+)\s*mS", probe_output)
            temp_match = re.search(r"([\d.]+)\s*°C", probe_output)
    
            if voltage_match and cond_match and temp_match:
                voltage_adc = float(voltage_match.group(1))
                conductivity_adc = float(cond_match.group(1))
                temp = float(temp_match.group(1))
    
                # Force all calibration values to float
                yintercept = float(yintercept)
                slope = float(slope)
                k_value = float(k_value)
    
                # Apply calibration math
                pH = round(((voltage_adc * 1000) - yintercept) / slope, 3)
                conductivity = round(conductivity_adc * k_value * 1000, 3)

    
                return {
                    'pH': pH,
                    'conductivity_mS-cm': conductivity,
                    'temperature_c': temp
                }
            else:
                print("Unexpected format in probe output.")
                return None
        except Exception as e:
            print(f"Error parsing readings: {e}")
            return None

    def get_raw_probe_values(self):
        probe_output = self.send_command('adc')
        print(f"Raw probe output: {probe_output}")
    
        try:
            voltage_match = re.search(r"([-+]?\d*\.?\d+)\s*V", probe_output)
            cond_match = re.search(r"([\d.]+)\s*mS", probe_output)
            temp_match = re.search(r"([\d.]+)\s*°C", probe_output)
    
            if voltage_match and cond_match and temp_match:
                voltage_adc = float(voltage_match.group(1))
                conductivity_adc = float(cond_match.group(1))
                temp = float(temp_match.group(1))
    
                return {
                    'voltage_V': voltage_adc,
                    'conductivity_mS': conductivity_adc,
                    'temperature_C': temp
                }
            else:
                print("Unexpected format in probe output.")
                return None
        except Exception as e:
            print(f"Error parsing readings: {e}")
            return None

    #-------Get ph equation values -----
    #Get slope and  y-intercepts 
    def get_ph_calibration_values(self):
        output = self.calibration_info_ph()
        match = re.search(r"CAL Channel 1 n = .*?Slope = ([\-\d.]+) V/pH, offset = ([\d.]+) V", output)
        
        if match:
            slope = float(match.group(1))
            offset = float(match.group(2))
            print("\nExtracted Calibration Values:")
            print(f"Slope: {slope} V/pH")
            print(f"Offset (Y-intercept): {offset} V")
            return slope, offset
        else:
            print("\nSlope and offset not found.")
            return None, None
        
#-----------------------------clear all inputs --------------------------------#
        
    def clear_all_inputs(self):
        return self.send_command('cal channel 1 remove all')
        

if __name__ == "__main__":
    port = 'COM4'  # Change to your actual COM port if different
    probe = SerialDevice(port)

    try:
        
        print(probe.clear_all_inputs())
        #----------Conductivity test ----------------#
        print(probe.set_conductivity_channel())
        print(probe.send_zero_conductivity())
        
        print(probe.send_range_conductivity(2))
        
        print(probe.send_set_conductivity(1.508))
        #----------Ph test ----------------#
        
        print(probe.set_ph_channel())
        print(probe.set_range_ph(200))
        print(probe.get_range_ph())
        print(probe.set_unit_ph())
        
        
        calibration_data = [
        (1, 4.00, 179.097),  # point 1 = pH 4.00 at 177.5 mV
        (2, 7.00, 6.056),    # point 2 = pH 7.00 at 0.0 mV
        (3, 10.00, -170.211) # point 3 = pH 10.00 at -177.5 mV
        ]

        for point_num, ph, mv in calibration_data:
            probe.set_calibration_ph(point_num, ph, mv)

                
        print(probe.calibration_info_ph())
        
        #---------- RTD test ----------------#
        print(probe.set_temp_channel())
        print(probe.set_range_temp(125))
        print(probe.get_range_temp())
        print(probe.get_ranges_temp())
        print(probe.get_reading_temp())
        
        
        #Setting temp probe to the appropriate channel (channel 3)
        print(probe.set_off_channel())
        
        
        
        
        
    finally:
        print(probe.clear_all_inputs())
        print("\nClosing connection...")
        probe.close()
