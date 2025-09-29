"""
Author: Molly Perez
Mentor: Lauren Valentino 

Purpose: 
    This program defines a Python class, `MasterFlex`, for controlling a peristaltic pump
    used in Capacitive Deionization (CDI) experiments via serial communication. The class 
    manages the connection to the pump through a specified COM port and includes methods 
    to send commands and retrieve data.
    
Methods:
    Key functionalities of the `MasterFlex` class include:
        Establishing a serial connection with the probe using parameters like baud rate, byte size, stop bits, and parity.
        retriving slope and y-intercept to be able to convert to correct speed
        Sending commands to the probe and receiving responses, with methods for specific operations such as:
            
        Pump Commands:
            -'send_command()': Added this send command to be able to decode the command sent to the MasterFlex Pump be able to understand correctly.
            -'initialize_pump()': Initialize the pump with correct satellite number hardcoded the satellite number to keep it consistent for every segment 
            -'send_enq_command()': The command is sent to send_command function make sure its encodeded correction
            -'set_speed_revolution()': This function sets the correct speed and revolutions for the specific satellite number.
            -'stop_pump()': Stop function makes the pump stop rotating and pumping 
            -'close()': Closes Serial Connection
            -'convert_ml_to_speed()': Converts the ml to the correct speed
            -'convert_and_set_speed()' Not in use now.
"""

import serial

class MasterFlex:
    def __init__(self, port,rpm_slope, intercept, baudrate=4800, timeout=1, bytesize=serial.SEVENBITS, stopbits=serial.STOPBITS_ONE, parity=serial.PARITY_ODD):
        self.serialPort = serial.Serial(
            port=port,
            baudrate=baudrate,
            timeout=timeout,
            bytesize=bytesize,
            stopbits=stopbits,
            parity=parity
        )
        
        self.rpm = rpm_slope
        self.b = intercept
        print(f"Using calibration slope = {self.rpm}, intercept = {self.b}")

        print(f"Serial port {port} opened successfully.")

    def send_command(self, command):
        self.serialPort.write(command.encode())
        print(f"Sent command: {command}")
        response = self.serialPort.read(10)
        print(f"Received response: {response}....")

    def initialize_pump(self, satellite_number):
        command = '\x02P01\r'
        self.send_command(command)
        print(f"Initialize command: {command}...")
        
    def send_enq_command(self):  
        enq_command = b'\x05'  
        self.serialPort.write(enq_command)  
        print(f"Sent <ENQ> command: {enq_command}")  
        response = self.serialPort.read(10)  
        print(f"Received response: {response}")      
    
    def set_speed_revolution(self, satellite_number, speed):
        revolutions = 99999
        command = f'\x02P{satellite_number:02d}S-{speed:06.1f}V{revolutions:08.2f}G\r'
        self.send_command(command)
        print(f"Set speed and revolution command: {command}")
        
    def stop_pump(self, satellite_number):
        stop_command = f'\x02P{satellite_number:02d}Z\r'.encode('utf-8')
        self.serialPort.write(stop_command)
        print(f"Sent stop command: {stop_command}")
    
    def close(self):
        self.serialPort.close()
        print("Serial port closed")
        
    def convert_ml_to_speed(self, ml):
        return (ml - self.b) / self.rpm
        
    def convert_and_set_speed(self, satellite_number, ml):
        speed = self.convert_ml_to_speed(ml)
        print(f"Converted {ml} mL to speed: {speed:.2f} rpm")
        self.set_speed_revolution(satellite_number, speed)
        

   
        




    

