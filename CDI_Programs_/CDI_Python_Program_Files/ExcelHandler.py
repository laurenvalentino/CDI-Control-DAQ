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
Author: Eric Vu, Molly Perez
Mentor: Lauren Valentino

Purpose:
    This program defines a Python class, `ExcelHandler`, which manages the interaction 
    between experimental CDI programs and Excel spreadsheets. The class automates 
    reading inputs, writing experimental data, handling COM port configurations, and 
    saving results into structured Excel workbooks. It supports multiple CDI program 
    configurations (Program1–Program4), each with different experimental setups such as 
    2-segment, 4-segment, multi-podvu, or dual power supply experiments.

Methods:
    Key functionalities of the `ExcelHandler` class include:

        Workbook Management:
            - `load_workbook()`: Opens the Excel template and loads program-specific sheets.
            - `create_wb()`: Generates a new workbook and assigns a unique filename.
            - `save_wb()`: Saves the workbook with the generated filename.
            - `open_excel()`: Makes the Excel application visible.
            - `generate_filename()`: Creates unique filenames with timestamps.
            - `delete_sheet()`: Removes unwanted sheets.
            - `clear_cells()`: Resets specific cell ranges to default.
            - `delete_unused_macros()`: Removes unnecessary VBA modules.
            - `delete_unwanted_buttons()`: Removes buttons except specified ones.

        Data Input/Output:
            - `read_header_value()`: Reads a specific header from the Inputs sheet.
            - `read_all_headers()`: Collects all defined headers and their values.
            - `read_cell()`: Reads any cell given row and column indices.
            - `read_com_ports()`: Retrieves and formats COM ports for probe and pump.
            - `psu_string()`: Retrieves the power supply USB string from Excel.
            - `read_pump_calibration()`: Loads calibration slope and intercept values.
            - `write_calibration_data()`: Saves calibration data to Excel.
            - `write_cycle_count()`: Writes the experiment cycle count.

        Experiment Data Logging:
            - `write_data_program1_2()`: Writes data (voltage, current, conductivity, etc.) 
              for Programs 1 and 2.
            - `write_data_program3()`: Logs data including pH and temperature for Program 3.
            - `write_data_program4()`: Handles dual power supply data logging for Program 4.
"""

import os
from datetime import datetime
import xlwings as xw

class ExcelHandler:
    
    HEADERS = {
            
    
            "Program1": {
                    # 2-Segment CDI  (Conductivity Only)
                "Reading interval (s)": ("Inputs", 3, 7),
                "Pre-run pump voltage (VDC)": ("Inputs", 4, 7),
                "Pre-run pump current (mA)": ("Inputs", 5, 7),
                "Pre-run pump time (s)": ("Inputs", 6, 7),
                "Pre-run pump Flow rate (mL/m)": ("Inputs", 7, 7),
    
                "Segment 1 voltage (VDC)": ("Inputs", 9, 7),
                "Segment 1 current (mA)": ("Inputs", 10, 7),
                "Segment 1 time (s)": ("Inputs", 11, 7),
                "Segment 1 Flow rate (mL/m)": ("Inputs", 12, 7),
    
                "Segment 2 voltage (VDC)": ("Inputs", 19, 7),
                "Segment 2 current (mA)": ("Inputs", 20, 7),
                "Segment 2 time (s)": ("Inputs", 21, 7),
                "Segment 2 Flow rate (mL/m)": ("Inputs", 22, 7),
    
                "Post-run pump voltage (VDC)": ("Inputs", 29, 7),
                "Post-run pump current (mA)": ("Inputs", 30, 7),
                "Post-run pump time (s)": ("Inputs", 31, 7),
                "Post-run pump Flow rate (mL/m)": ("Inputs", 32, 7),
                
                
            },

            "Program2": {
                # 4-Segment CDI (2 intersegments) (Conductivity measurment only)
                "Reading interval (s)": ("Inputs", 3, 10),
                "Pre-run pump voltage (VDC)": ("Inputs", 4, 10),
                "Pre-run pump current (mA)": ("Inputs", 5, 10),
                "Pre-run pump time (s)": ("Inputs", 6, 10),
                "Pre-run pump Flow rate (mL/m)": ("Inputs", 7, 10),
    
                "Segment 1 voltage (VDC)": ("Inputs", 9, 10),
                "Segment 1 current (mA)": ("Inputs", 10, 10),
                "Segment 1 time (s)": ("Inputs", 11, 10),
                "Segment 1 Flow rate (mL/m)": ("Inputs", 12, 10),
                
                "Inter-segment 1 voltage (VDC)": ("Inputs", 14, 10),
                "Inter-segment 1 current (mA)": ("Inputs", 15, 10),
                "Intersegment 1 time (s)": ("Inputs", 16, 10),
                "Intersegment 1 flow rate (mL/m)": ("Inputs", 17, 10),
    
                "Segment 2 voltage (VDC)": ("Inputs", 19, 10),
                "Segment 2 current (mA)": ("Inputs", 20, 10),
                "Segment 2 time (s)": ("Inputs", 21, 10),
                "Segment 2 Flow rate (mL/m)": ("Inputs", 22, 10),
                
                "Inter-segment 2 voltage (VDC)": ("Inputs", 24, 10),
                "Inter-segment 2 current (mA)": ("Inputs", 25, 10),
                "Intersegment 2 time (s)": ("Inputs", 26, 10),
                "Intersegment 2 flow rate (mL/m)": ("Inputs", 27, 10),
    
                "Post-run pump voltage (VDC)": ("Inputs", 29, 10),
                "Post-run pump current (mA)": ("Inputs", 30, 10),
                "Post-run pump time (s)": ("Inputs", 31, 10),
                "Post-run pump Flow rate (mL/m)": ("Inputs", 32, 10),
            },
            
            "Program3": {
                # 2-Segment CDI (Conductivity, PH & Temperature Measurment)
                "Reading interval (s)": ("Inputs", 3, 13),
                "Pre-run pump voltage (VDC)": ("Inputs", 4, 13),
                "Pre-run pump current (mA)": ("Inputs", 5, 13),
                "Pre-run pump time (s)": ("Inputs", 6, 13),
                "Pre-run pump Flow rate (mL/m)": ("Inputs", 7, 13),
    
                "Segment 1 voltage (VDC)": ("Inputs", 9, 13),
                "Segment 1 current (mA)": ("Inputs", 10, 13),
                "Segment 1 time (s)": ("Inputs", 11, 13),
                "Segment 1 Flow rate (mL/m)": ("Inputs", 12, 13),
    
                "Segment 2 voltage (VDC)": ("Inputs", 19, 13),
                "Segment 2 current (mA)": ("Inputs", 20, 13),
                "Segment 2 time (s)": ("Inputs", 21, 13),
                "Segment 2 Flow rate (mL/m)": ("Inputs", 22, 13),
    
                "Post-run pump voltage (VDC)": ("Inputs", 29, 13),
                "Post-run pump current (mA)": ("Inputs", 30, 13),
                "Post-run pump time (s)": ("Inputs", 31, 13),
                "Post-run pump Flow rate (mL/m)": ("Inputs", 32, 13),
            },
            
            "Program4": {
                # 2-Segment CDI with 2 Power Supplies               
                "Reading interval (s)": ("Inputs", 3, 16),
                "Pre-run pump voltage (VDC)": ("Inputs", 4, 16),
                "Pre-run pump current (mA)": ("Inputs", 5, 16),
                "Pre-run pump time (s)": ("Inputs", 6, 16),
                "Pre-run pump Flow rate (mL/m)": ("Inputs", 7, 16),
                
                "Pre-run pump voltage channel 2 (VDC)": ("Inputs", 4, 19),
                "Pre-run pump current channel 2 (mA)": ("Inputs", 5, 19),
    
                "Segment 1 voltage (VDC)": ("Inputs", 9, 16),
                "Segment 1 current (mA)": ("Inputs", 10, 16),
                "Segment 1 time (s)": ("Inputs", 11, 16),
                "Segment 1 Flow rate (mL/m)": ("Inputs", 12, 16),
                
                "Segment 1 voltage channel 2 (VDC))": ("Inputs", 9, 19),
                "Segment 1 current channel 2 (mA)": ("Inputs", 10, 19),
    
                "Segment 2 voltage (VDC)": ("Inputs", 19, 16),
                "Segment 2 current (mA)": ("Inputs", 20, 16),
                "Segment 2 time (s)": ("Inputs", 21, 16),
                "Segment 2 Flow rate (mL/m)": ("Inputs", 22, 16),
                
                "Segment 2 voltage channel 2 (VDC)": ("Inputs", 19, 19),
                "Segment 2 current channel 2 (mA)": ("Inputs", 20, 19),
    
                "Post-run pump voltage (VDC)": ("Inputs", 29, 16),
                "Post-run pump current (mA)": ("Inputs", 30, 16),
                "Post-run pump time (s)": ("Inputs", 31, 16),
                "Post-run pump Flow rate (mL/m)": ("Inputs", 32, 16),
                
                "Post-run pump voltage channel 2 (VDC)": ("Inputs", 29, 19),
                "Post-run pump current channel 2 (mA)": ("Inputs", 30, 19),
            }
            
        }
            
    # Shared for all experiments
    DATE_COLUMN = 4
    TIME_COLUMN = 5
    ELAPSED_TIME_COLUMN = 6 
    VOLTAGE_COLUMN = 7
    CURRENT_COLUMN = 8
    
    # Varies by experiment program 1,2 &3
    CONDUCTIVITY_COLUMN = 9 #For Experiment program 1,2 &3
    
    # Only for experiemnt 3
    PH_PROBE = 10
    TEMP_PROBE =11
    
    #only for experiment 1&2
    CYCLE_COUNT_COLUMN = 10
    SEGMENT_COLUMN = 11
    FLOWRATE_COLUMN = 12
    
    #Onlt for program 3&4
    CYCLE_COUNT_COLUMN_2 = 12
    SEGMENT_COLUMN_2 = 13
    FLOWRATE_COLUMN_2 = 14
    
    #Only for Program 4
    VOLTAGE_COLUMN_2 = 9
    CURRENT_COLUMN_2 = 10
    CONDUCTIVITY_COLUMN_2 = 11   # for Experiment program 4
    

    def __init__(self, template_file, program_name):
        self.template_file = template_file
        self.program_name = program_name  # e.g., "Program1"
        self.filename = None
        self.current_row = 2
        self.load_workbook()

    def load_workbook(self):
        try:
            self.wb = xw.Book(self.template_file)
            self.wb.app.visible = False
    
            self.input_sheet = self.wb.sheets['Inputs']
            self.data_sheet = self.wb.sheets[self.program_name]  # dynamic sheet assignment
        except FileNotFoundError:
            raise FileNotFoundError(f"Template file {self.template_file} not found.")
        except Exception as e:
            raise RuntimeError(f"Error loading workbook: {e}")
            
    def create_wb(self):
        self.generate_filename()
        self.open_excel()  # Open the newly created workbook

    def read_header_value(self, label):
        if label not in self.HEADER_LOCATIONS:
            raise ValueError(f"Unknown header label: {label}")
        sheet_name, row, col = self.HEADER_LOCATIONS[label]
        sheet = self.wb.sheets[sheet_name]
        return sheet.range((row, col)).value

    def read_all_headers(self):
        return {header: self.read_header_value(header) for header in self.HEADERS}

    def write_data_program1_2(self, elapsed_time, voltage, current, conductivity, segment, flowrate, cycle_count):
        current_date = datetime.now().strftime("%y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")
    
        row_data = [
            current_date,              # DATE_COLUMN = 4
            current_time,              # TIME_COLUMN = 5
            elapsed_time,              # ELAPSED_TIME_COLUMN = 6
            voltage,                   # VOLTAGE_COLUMN = 7
            current,                   # CURRENT_COLUMN = 8
            conductivity,              # CONDUCTIVITY_COLUMN = 9
            cycle_count,               # CYCLE_COUNT_COLUMN = 10
            segment,                   # SEGMENT_COLUMN = 11
            flowrate                   # FLOWRATE_COLUMN = 12
        ]
    
        self.data_sheet.range((self.current_row, self.DATE_COLUMN), 
                              (self.current_row, self.FLOWRATE_COLUMN)).value = row_data
        self.current_row += 1

    def write_data_program3(self, elapsed_time, voltage, current, conductivity, pH, temp_C, segment, flowrate, cycle_count):
        """Write real-time data into the Excel file."""
        current_date = datetime.now().strftime("%y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")

        # Prepare all data for a single-row update
        row_data = [
            current_date,   # 4
            current_time,   # 5
            elapsed_time,   # 6
            voltage,        # 7
            current,        # 8
            conductivity,    # 9
            pH,             # PH_PROBE = 10
            temp_C,          # TEMP_PROBE = 11
            cycle_count,  # CYCLE_COUNT_COLUMN_2 = 12
            segment,        # SEGMENT_COLUMN_2 = 13
            flowrate    # FLOWRATE_COLUMN_2 = 14
        ]

        # Update the entire row in one go
        self.data_sheet.range((self.current_row, self.DATE_COLUMN), 
                    (self.current_row, self.FLOWRATE_COLUMN_2)).value = row_data

        self.current_row += 1  # Move to the next row for the next update
        
    def write_data_program4(self, elapsed_time, voltage, current, voltage2, current2, conductivity, segment, flowrate, cycle_count):
        current_date = datetime.now().strftime("%y-%m-%d")
        current_time = datetime.now().strftime("%H:%M:%S")
    
        row_data = [
            current_date,                  # DATE_COLUMN = 4
            current_time,                  # TIME_COLUMN = 5
            elapsed_time,                  # ELAPSED_TIME_COLUMN = 6
            voltage,                       #voltage 7
            current,                       # Voltage 8
            voltage2,                      # VOLTAGE_COLUMN_2 = 9 
            current2,                      # CURRENT_COLUMN_2 = 10
            conductivity,                  # CONDUCTIVITY_COLUMN_2 = 11
            cycle_count,                   # CYCLE_COUNT_COLUMN_2 = 12
            segment,                       # SEGMENT_COLUMN_2 = 13
            flowrate                       # FLOWRATE_COLUMN_2 = 14
        ]
    
        self.data_sheet.range((self.current_row, self.DATE_COLUMN),
                              (self.current_row, self.FLOWRATE_COLUMN_2)).value = row_data
        self.current_row += 1

#--------------- Cycle_count, Calibration_data, COM_Ports ---------------------#     
    def write_cycle_count(self, cycle_count):
        "write cycle count to cell C30"
        self.input_sheet.range("C30").value = cycle_count
        
    def write_calibration_data(self, slope, offset):
        "Write slope and y_intercept"
        self.input_sheet.range("C24").value = slope
        self.input_sheet.range("C25").value = offset
            
    def read_com_ports(self):
            """
            Reads the COM port *numbers* for probe and pump from Excel
            and formats them as COM strings (e.g., 'COM4').
            Returns: (probe_port_str, pump_port_str)
            """
            try:
                probe_number = int(self.read_cell(2, 2))  # Example cell: Probe COM number
                pump_number = int(self.read_cell(3, 2))   # Example cell: Pump COM number
                
                probe_port = f"COM{probe_number}"
                pump_port = f"COM{pump_number}"
                
                return probe_port, pump_port
            except Exception as e:
                print(f"Error reading COM port numbers from Excel: {e}")
                return None, None      
            
    def psu_string(self):
            """
            Reads the powersupply Serial number
            """
            try:
                serial_number = str(self.read_cell(7, 2))  # Example cell: Probe COM number
                
                power_supply_resource = f"USB0::6833::3601::{serial_number}::0::INSTR"
                
                return power_supply_resource
            except Exception as e:
                print(f"Error reading Powersupply string from Excel: {e}")
                return None   

    def read_pump_calibration(self):
        """Read slope and y-intercept from Inputs sheet for calibration."""
        rpm_slope = self.input_sheet.range("C10").value
        pump_intercept = self.input_sheet.range("C11").value
        
        if rpm_slope is None or pump_intercept is None:
            raise ValueError("Missing calibration values in Excel (C10/C11).")

        
        return rpm_slope, pump_intercept
    
#--------------- Cycle_count, Calibration_data, COM_Ports ---------------------#     
    
    def read_cell(self, row, col):
        """Read the value of a cell given its row and column indices (1-based)."""
        if row < 1 or col < 1:
            raise ValueError("Row and column indices must be 1 or greater.")
        
        return self.input_sheet.range((row, col)).value
    
    def generate_filename(self):
        current_date = datetime.now().strftime("%Y%m%d")
        base_filename = f"{current_date} CDI.xlsm"
        counter = 1

        # Ensure the file saves in the current working directory
        current_directory = os.getcwd()  # Get the current working directory
        filename = os.path.join(current_directory, base_filename)

        # Check if the file already exists and if so, find a unique filename
        while os.path.exists(filename):
            filename = os.path.join(current_directory, f"{current_date} CDI ({counter}).xlsm")
            counter += 1

        self.filename = filename
        self.wb.save(self.filename)

    def save_wb(self):
        """Save the workbook to the generated filename."""
        if self.wb:
            try:
                self.wb.save(self.filename)
                print(f"Excel file saved: {self.filename}")
            except Exception as e:
                raise RuntimeError(f"Error saving workbook: {e}")
        else:
            raise RuntimeError("Workbook has not been created. Call create_wb() first.")

    def open_excel(self):
        """Make the Excel application visible to view the file."""
        if self.wb:
            self.wb.app.visible = True
            self.wb.app.activate()
        else:
            raise RuntimeError("Workbook has not been created. Call create_wb() first.")

    def delete_sheet(self, sheet_name):
        """
        Deletes a sheet from the workbook if it exists.
        """
        try:
            sheet = self.wb.sheets[sheet_name]
            sheet.delete()
            print(f"Sheet '{sheet_name}' deleted.")
        except Exception as e:
            print(f"Could not delete sheet '{sheet_name}': {e}")

    def clear_cells(self, sheet_name, cell_range):
        """
        Clears the contents, background, and borders of the specified cell range
        in the given sheet, and resets formatting to default (white fill, no borders).
        """
        try:
            sheet = self.wb.sheets[sheet_name]  # Access the specified sheet
            target_range = sheet.range(cell_range)
    
            # Unmerge any merged cells in the range
            target_range.api.UnMerge()
    
            # Clear contents
            target_range.value = None
    
            # Set white background (RGB for white)
            target_range.color = (255, 255, 255)
    
            # Remove all borders (Excel border IDs 7 to 12)
            for border_id in range(7, 13):
                border = target_range.api.Borders(border_id)
                border.LineStyle = 0  # xlLineStyleNone
    
            print(f"Cleared and formatted cells '{cell_range}' in sheet '{sheet_name}'.")
        except Exception as e:
            print(f"Error clearing cells '{cell_range}' in sheet '{sheet_name}': {e}")
    
    def delete_unused_macros(self, keep_modules=None):
        """
        Deletes VBA modules that are not in the keep_modules list.
        Only works if workbook has macros (xlsm).
        """
        try:
            if keep_modules is None:
                keep_modules = []
    
            vb_project = self.wb.api.VBProject
            for i in reversed(range(vb_project.VBComponents.Count)):
                component = vb_project.VBComponents.Item(i + 1)
                name = component.Name
                if name not in keep_modules and component.Type == 1:  # 1 = standard module
                    vb_project.VBComponents.Remove(component)
                    print(f"Deleted macro module: {name}")
        except Exception as e:
            print(f"Error deleting macro modules: {e}")
                      
    def delete_unwanted_buttons(self, sheet_name, keep_buttons=None):
        """
        Deletes all Form Control buttons on a specified sheet except those in keep_buttons.
        Does NOT save or close the workbook.
        """
        if keep_buttons is None:
            keep_buttons = []
    
        ws = self.wb.sheets[sheet_name]
    
        for shape in ws.api.Shapes:
            if shape.Type == 8:  # 8 = Form Control Button
                if shape.Name not in keep_buttons:
                    shape.Delete()
        print(f"Deleted unwanted buttons in sheet '{sheet_name}' except {keep_buttons}")

            
