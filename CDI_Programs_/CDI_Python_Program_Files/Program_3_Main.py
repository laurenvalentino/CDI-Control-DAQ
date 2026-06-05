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
Author: Molly Perez & Eric Vu
Mentor: Lauren Valentino

Purpose:
    This Python script automates the CDI experiment by controlling a conductivity probe, pump, and power supply. 
    It also collects and records experimental data in an excel.xls file.
    This is Program Version 3, which operates with 2 segment per cycles, along with a pre-run and post-run. 
    The experiment cycles run continuously for the number of cycles specified by the user, with user-defined times for each cycle. 
    If the stop macro is triggered, the script allows the current cycle to complete before proceeding to the post-run.
    
Experiment Sequence:
    Pre-Run
    Segment 1
    Segment 2
    Post-Run
    

For each part of the experiment, the script uses user input from the Excel interface to set Voltage, Current, Time, and Flowrate.

Data Collection: 
    date, time, elapse time, voltage (V), Current (mA), Conductivity (uS/cm), pH, Temperature, cycle, segment, and Flowrate

How is this program different?
    This program differs as it uses the Quad Multi Function isoPod that allows data collection for 2 extra parameters the pH and Temperature. 
    
    Channel 1: PH probe 
    Channel 2: Conductivity Probe
    Channel 3: Temperature Probe
"""
import multi_edaq_probe_working as probe
from ExcelHandler import ExcelHandler
import pump_initialize as pump
import psu
import threading
import time
import os
import sys
import subprocess
import psutil



SATELLITE_NUMBER = 1
#Name of excel sheet
EXISTING_FILE_TEMPLATE = 'CDI_Program_.xlsm'
STOP_SIGNAL_FILE = 'stop_signal.txt'  # File to monitor for the stop signal



def close_open_terminals():
    """
    Closes all open terminal windows except the current one running this script.
    This ensures that no other terminal sessions interfere with the experiment.
    """
    try:
        current_pid = os.getpid()  # Get the PID of the current process

        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] in ['cmd.exe', 'powershell.exe'] and proc.info['pid'] != current_pid:
                # Close only other terminals, not the one running this script
                proc.terminate()

        print("Closed all other terminal windows except the current one.")
    except Exception as e:
        print(f"An error occurred while closing terminals: {e}")

def monitor_file(stop_event):
    """
    Monitors a file (stop_signal.txt) for a 'stop' signal. When detected,
    the stop_event is triggered to stop the experiment. This stop event is triggered once the stop macro is clicked in excel.
    """
    while not stop_event.is_set():
        if os.path.exists(STOP_SIGNAL_FILE):
            with open(STOP_SIGNAL_FILE, 'r') as f:
                signal = f.read().strip()
                if signal.lower() == 'stop':
                    stop_event.set()  # Trigger the stop event if 'stop' is found in the file
        time.sleep(1)  # Check the file every 1 second

def stop_experiment(pump_device=None, power_supply=None, probe_device=None, excel_handler=None, satellite_number=None):
    """
    Safely stops the experiment by closing all devices and saving the Excel workbook.
    This function ensures all resources are properly released.
    """
    print("Stopping the experiment...")

    if pump_device:
        pump_device.stop_pump(satellite_number)
        pump_device.close()

    if power_supply:
        power_supply.close()

    if probe_device:
        # clear all data of Ph inputs
        probe_device.clear_all_inputs()
        probe_device.close()

    if excel_handler:
        excel_handler.save_wb()

def initialize_excel_handler():
    """
    Initializes the ExcelHandler class to handle data logging in the experiment.
    The handler is responsible for interacting with the workbook template.
    """
    print("Initializing Excel Handler...")
    excel_handler = ExcelHandler(EXISTING_FILE_TEMPLATE)
    excel_handler.create_wb()  # Create or open the workbook
    print("Excel Handler initialized.")
    return excel_handler

def initialize_power_supply(excel_handler, power_supply_resource):
    """
    Initializes the power supply and sets the initial voltage and current 
    based on values read from the Excel sheet (pre-run settings).
    """
    print("Initializing Power Supply...")
    power_supply = psu.psu(power_supply_resource)

    # Set initial voltage and current from pre-run settings in Excel
    pre_run_voltage = float(excel_handler.read_cell(4, 13))
    pre_run_current = float(excel_handler.read_cell(5, 13))
    power_supply.set_voltage_current(pre_run_voltage, pre_run_current)
    print("Power Supply initialized with voltage:",
          pre_run_voltage, "and current:", pre_run_current)

    return power_supply, pre_run_voltage, pre_run_current

def initialize_probe(excel_handler, probe_port):
    """
    Initializes the probe for measuring pH, conductivity, and Temp. Sends a series of 
    commands to configure the probe before starting the experiment, including calibration for conducitvity using the K Value 
    in the user input in excel, and the calibration (pH , V )for the pH probe calibration. 
    """
    print("Initializing Multi Quad Probe Conductivity ...")
    
    try:
        probe_device = probe.SerialDevice(probe_port)
        probe_device.clear_all_inputs()

        # ------- Conductivity Probe Setup -------
        probe_device.set_conductivity_channel()
        probe_device.send_zero_conductivity()
        probe_device.send_range_conductivity(2)

        # Read and convert k_value with error handling
        try:
            k_value_raw = excel_handler.read_cell(15, 2)
            print(f"Read k_value: {k_value_raw} (type: {type(k_value_raw)})")
            k_value = float(k_value_raw)
            print(f"Converted k_value: {k_value} (type: {type(k_value)})")
        except Exception as e:
            print(f"Error reading or converting k_value: {e}")
            return None, None, None, None  # Gracefully abort probe initialization

        # Send k_value to probe
        try:
            probe_device.send_set_conductivity(k_value)
            print("k_value successfully sent to probe.") # Printed when succesful 
        except Exception as e:
            print(f"Error while sending k_value to probe: {e}")
            return None, None, None, None

        # ------- pH Calibration Setup -------
        print("Initializing Multi Quad Probe pH ...")
        probe_device.set_ph_channel()
        probe_device.set_range_ph(200)
        probe_device.get_range_ph()
        probe_device.set_unit_ph()

        calibration_data = [
            (1, float(excel_handler.read_cell(20, 2)), float(excel_handler.read_cell(20, 3))),
            (2, float(excel_handler.read_cell(21, 2)), float(excel_handler.read_cell(21, 3))),
            (3, float(excel_handler.read_cell(22, 2)), float(excel_handler.read_cell(22, 3)))
        ]

        for point_num, ph, mv in calibration_data:
            probe_device.set_calibration_ph(point_num, ph, mv)

        probe_device.calibration_info_ph()
        slope, offset = probe_device.get_ph_calibration_values()

        if slope is not None and offset is not None:
            excel_handler.write_calibration_data(slope, offset)

        # ------- RTD Temperature Setup -------
        print("Initializing Multi Quad Probe Temp ...")
        probe_device.set_temp_channel()
        probe_device.set_range_temp(125)
        probe_device.get_range_temp()
        probe_device.get_ranges_temp()
        probe_device.get_reading_temp()

        # Turn off Channel 4
        probe_device.set_off_channel()

        print("Probe initialized.") 
        return probe_device, slope, offset, k_value

    except Exception as e:
        print(f"Probe initialization failed: {e}")
        return None, None, None, None

def initialize_pump(excel_handler, pump_port):
    """
    Initializes the pump by connecting to it via the specified port.
    A temporary fix involves connecting twice due to an initialization issue.
    As well this method sends the power supply calibration to the pump decive class to be able to 
    convert the flowrate to the correct speed (rpm) for the pump. 
    """
    print("Initializing Pump...")

    rpm_slope, intercept = excel_handler.read_pump_calibration()
    
    print(f"Slope: {rpm_slope} and y-int: {intercept}")
    
    # Connect and initialize pump
    pump_device = pump.MasterFlex(pump_port, rpm_slope, intercept)
    pump_device.initialize_pump(SATELLITE_NUMBER)
    pump_device.send_enq_command()

    # Temporary fix to execute command, requires connecting twice
    pump_device.close()
    pump_device = pump.MasterFlex(pump_port, rpm_slope, intercept)
    pump_device.initialize_pump(SATELLITE_NUMBER)
    pump_device.send_enq_command()

    print("Pump initialized.")
    return pump_device, SATELLITE_NUMBER

def read_experiment_parameters(excel_handler):
    """
    Reads and returns all experiment parameters from the Excel sheet, storing them 
    in a dictionary for easy access during the experiment.
    
    All the inputs are hardcoded based on the excel set up this is for program 3 user inputs. Column 13
    
    WARNING: Changing a cell in excel will raise an error and it will not allow to run the program via excel. 
    """
    
    print("Initializing experimental Parameters...")
    
    parameters = {
        "reading_interval": float(excel_handler.read_cell(3, 13)),
        "pre_run_voltage": float(excel_handler.read_cell(4, 13)),
        "pre_run_current": float(excel_handler.read_cell(5, 13)),
        "pre_run_time": float(excel_handler.read_cell(6, 13)),
        "pre_run_flowrate": float(excel_handler.read_cell(7, 13)),

        "segment1_voltage": float(excel_handler.read_cell(9, 13)),
        "segment1_current": float(excel_handler.read_cell(10, 13)),
        "segment1_time": float(excel_handler.read_cell(11, 13)),
        "segment1_flowrate": float(excel_handler.read_cell(12, 13)),

        "segment2_voltage": float(excel_handler.read_cell(19, 13)),
        "segment2_current": float(excel_handler.read_cell(20, 13)),
        "segment2_time": float(excel_handler.read_cell(21, 13)),
        "segment2_flowrate": float(excel_handler.read_cell(22, 13)),

        "post_run_voltage": float(excel_handler.read_cell(29, 13)),
        "post_run_current": float(excel_handler.read_cell(30, 13)),
        "post_run_time": float(excel_handler.read_cell(31, 13)),
        "post_run_flowrate": float(excel_handler.read_cell(32, 13)),

        "number_of_cycles": float(excel_handler.read_cell(29, 3)),
    }
            

    return parameters

def collect_and_log_data(excel_handler, power_supply, probe_device, overall_start_time, segment_name, flowrate, reading_interval, last_logged_time, cycle_count, slope, yintercept, k_value):
    """Collects and logs data from devices to the Excel file."""

    current_time = time.monotonic()
    elapsed_since_last_log = current_time - last_logged_time

    if elapsed_since_last_log >= reading_interval:
        elapsed_time = round(current_time - overall_start_time, 1)
        
        #Using method to aquire voltage and current readings
        voltage = power_supply.measure_voltage()
        current = power_supply.measure_current()

        # Get raw values from the probe
        raw_data = probe_device.get_raw_probe_values()
        if raw_data is None:
            print("Skipping logging due to malformed probe output.")
            return last_logged_time

        try:
            voltage_adc = float(raw_data['voltage_V'])
            conductivity_adc = float(raw_data['conductivity_mS'])
            temp_C = float(raw_data['temperature_C'])

            # Ensure calibration constants are float
            slope = float(slope)
            yintercept = float(yintercept)
            k_value = float(k_value)

            # Apply calibration math to obtain correct conversion for data collection
            ph = round(((voltage_adc * 1000) - yintercept) / slope, 3)
            conductivity = round(conductivity_adc * k_value * 1000, 3)

            # Write data to Excel sending to the correct method as the way the data in collected and written
            excel_handler.write_data_program3(
                elapsed_time, voltage, current, conductivity, ph, temp_C, segment_name, flowrate, cycle_count)

        except Exception as e:
            print(f"Error during calibration math or data logging: {e}")

        return current_time
    return last_logged_time

def clear_all_unwanted(excel_handler):
    """
    Cleans up the Excel workbook by removing unused data, buttons, 
    and extra sheets to ensure only relevant information remains for the specific program being used. 

    """
    print("Clearing cells and borders.")
    ranges_to_clear = [
        "F1:F38", "G1:G38",
        "I1:I38", "J1:J38",
        "O1:O38", "P1:P38",
        "Q1:Q38", "R1:R38",
        "S1:S38"
    ]
    for cell_range in ranges_to_clear:
        excel_handler.clear_cells("Inputs", cell_range)

   # print("Clearing macros")
    #excel_handler.delete_unused_macros(keep_modules=["Module8", "Module7", "Module9"]) #keeps only these modules 

    print("Clearing buttons")
    excel_handler.delete_unwanted_buttons("Inputs", keep_buttons=["Button 21", "Button 20", "Button 22"])  #Keeps only these buttons 

    print("Deleting sheets")
    sheets_to_delete = ["Program_1", "Program_2", "Program_4"] #Deletes these sheets
    for sheet_name in sheets_to_delete:
        excel_handler.delete_sheet(sheet_name)

def main():
    """
    Runs the CDI experiment from start to finish for Program 3

    - Initializes the Excel workbook and clears unwanted data.
    - Connects to the power supply, pump, and probe.
    - Reads experiment parameters from the workbook.
    - Executes Pre-Run, repeated Segment cycles, and Post-Run.
    - Logs data at the configured interval and updates cycle counts.
    - Monitors a stop signal for safe termination.
    - Cleans up devices, threads, and files after completion or error.
    """
    close_open_terminals()

    if os.path.exists(STOP_SIGNAL_FILE):
        os.remove(STOP_SIGNAL_FILE)
        print(f"{STOP_SIGNAL_FILE} from a previous experiment has been removed.")

    if len(sys.argv) != 2:
        print("Usage: python main.py <workbook_path>")
        sys.exit(1)
        
    workbook_path = sys.argv[1]

    # Initialize placeholders for safety in case of exception

    stop_event = None
    file_thread = None
    satellite_number = None

    try:
        excel_handler = ExcelHandler(workbook_path, "Program_3")
        
        #Clearning certain cells in inputs tab
        print("Clearing up excel....")
        clear_all_unwanted(excel_handler)
                
        excel_handler.create_wb()
        print("Excel Handler initialized.")
        
        probe_port, pump_port = excel_handler.read_com_ports()

        if not probe_port or not pump_port:
            print("Error: COM ports could not be read.")
            sys.exit(1)
            
        print("COM port read...")

        
        power_supply_resource = excel_handler.psu_string()
        if not power_supply_resource:
            print("Error: power_Supply string could not be read.")
            sys.exit(1)
            
        print("Power Supply string read...")
        

        power_supply, segment1_voltage, segment1_current = initialize_power_supply(excel_handler, power_supply_resource)
        probe_device, slope, yintercept, k_value = initialize_probe(excel_handler, probe_port)
        pump_device, satellite_number = initialize_pump(excel_handler, pump_port)
        experiment_parameters = read_experiment_parameters(excel_handler)

        print("Starting experiment...")

        # --- Pre-Run --
        pre_run_voltage = experiment_parameters["pre_run_voltage"]
        pre_run_current = experiment_parameters["pre_run_current"]
        pre_run_flowrate = experiment_parameters["pre_run_flowrate"]
        pre_run_time = experiment_parameters["pre_run_time"]

        print("Starting Pre-Run...")
        power_supply.set_voltage_current(pre_run_voltage, pre_run_current)
        pump_device.convert_and_set_speed(satellite_number, pre_run_flowrate)

        # Record overall start time using monotonic time
        overall_start_time = time.monotonic()
        pre_run_start_time = time.monotonic()
        last_logged_time = overall_start_time

        while time.monotonic() - pre_run_start_time < pre_run_time:
            last_logged_time = collect_and_log_data(excel_handler, power_supply, probe_device, overall_start_time, "Pre-Run",
                                                    pre_run_flowrate, experiment_parameters["reading_interval"], last_logged_time, 0, slope, yintercept, k_value)

        # --- Setup ---
        stop_event = threading.Event()
        file_thread = threading.Thread(target=monitor_file, args=(stop_event,))
        file_thread.start()

        num_cycles = int(experiment_parameters["number_of_cycles"])
        cycle_count = 0
        

        segments = [
            ("Segment 1",
             experiment_parameters["segment1_voltage"], 
             experiment_parameters["segment1_current"],
             experiment_parameters["segment1_flowrate"],
             experiment_parameters["segment1_time"]),
            
            ("Segment 2", 
             experiment_parameters["segment2_voltage"], 
             experiment_parameters["segment2_current"],
             experiment_parameters["segment2_flowrate"], 
             experiment_parameters["segment2_time"]),
        ]
        stop_requested = False
        cycle_count = 0

        # --- Main Loop ---
        while True:
            print(f"Cycle Count: {cycle_count}, Stop Requested: {stop_requested}, Num Cycles: {num_cycles}")
           
            if stop_requested or cycle_count >= num_cycles:
                print(" Experiment finished or stop request set")
                break
           
            print(f"Starting Cycle {cycle_count + 1} of {num_cycles}...")

            for segment, voltage, current, flowrate, duration in segments:
                print(f"Starting {segment}...")
                segment_start_time = time.monotonic()
                power_supply.set_voltage_current(voltage, current)
                pump_device.convert_and_set_speed(satellite_number, flowrate)

                if duration > 0:
                    while time.monotonic() - segment_start_time < duration:
                        last_logged_time = collect_and_log_data(
                            excel_handler, power_supply, probe_device, overall_start_time,
                            segment, flowrate, experiment_parameters["reading_interval"],
                            last_logged_time, cycle_count + 1, slope, yintercept, k_value
                        )

                        if stop_event.is_set():
                            stop_requested = True
                           

                if stop_requested and segment == "Segment 2":
                    print("Stop requested during segment. Exiting cycle early.")
                    break

            cycle_count += 1
            excel_handler.write_cycle_count(cycle_count)
            print(f"Cycle {cycle_count} completed.")

        # --- Post-Run ---
        print("All cycles complete or stop requested. Starting Post-Run...")
        print("Starting Post-Run...")
        post_run_voltage = experiment_parameters["post_run_voltage"]
        post_run_current = experiment_parameters["post_run_current"]
        post_run_flowrate = experiment_parameters["post_run_flowrate"]
        post_run_time = experiment_parameters["post_run_time"]
        power_supply.set_voltage_current(post_run_voltage, post_run_current)
        pump_device.convert_and_set_speed(satellite_number, post_run_flowrate)
        post_run_start_time = time.monotonic()


        post_run_start_time = time.monotonic()
        while time.monotonic() - post_run_start_time < post_run_time:
            last_logged_time = collect_and_log_data(
                excel_handler, power_supply, probe_device, overall_start_time, "Post-Run",
                post_run_flowrate, experiment_parameters["reading_interval"],
                last_logged_time, 0, slope, yintercept, k_value
            )

        print("Experiment completed.")
        excel_handler.save_wb()
        

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        stop_experiment(pump_device, power_supply, probe_device, excel_handler, satellite_number)
        stop_event.set()
        file_thread.join()
        if os.path.exists(STOP_SIGNAL_FILE):
            os.remove(STOP_SIGNAL_FILE)



if __name__ == '__main__':
    main()
