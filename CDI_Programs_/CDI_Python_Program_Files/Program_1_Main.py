"""
Author: Molly Perez & Eric Vu
Mentor: Lauren Valentino

Purpose:
    This Python script automates the CDI experiment by controlling a conductivity probe, pump, and power supply. 
    It also collects and records experimental data in an excel.xls file.
    This is Program Version 1, which operates with two segment per cycle, along with a pre-run and post-run. 
    The experiment cycles run continuously for the number of cycles specified by the user, with user-defined times for each cycle. 
    If the stop macro is triggered, the script allows the current cycle to complete before proceeding to the post-run.
    
Experiment Sequence:
    Pre-Run
    Segment 1
    Segment 2
    Post-Run

For each part of the experiment, the script uses user input from the Excel interface to set Voltage, Current, Time, and Flowrate.

Data Collection:
    Date, time, elapse time, voltage (V), Current (mA), Conductivity (uS/cm), cycle, segment, and Flowrate

How is this program different?
    This program differs as it is only a 2 segment cycle. 
"""

import probe_working as probe
from ExcelHandler import ExcelHandler 
import pump_initialize as pump
import psu
import threading
import time
import os
import sys
import subprocess
import psutil


# Configuration constants
SATELLITE_NUMBER = 1
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
    pre_run_voltage = float(excel_handler.read_cell(4, 7))
    pre_run_current = float(excel_handler.read_cell(5, 7))
    power_supply.set_voltage_current(pre_run_voltage, pre_run_current)
    print("Power Supply initialized with voltage:", pre_run_voltage, "and current:", pre_run_current)

    return power_supply, pre_run_voltage, pre_run_current

def initialize_probe(excel_handler, probe_port):
    """
    Initializes the probe for measuring conductivity, sending a series of 
    commands to configure the probe before starting the experiment.
    """
    print("Initializing Probe...")

    # Initialize the probe device
    probe_device = probe.SerialDevice(probe_port)

    probe_device.send_zero()
    probe_device.send_range(2)  # Set measurement range    

    # Read k value from Excel
    try:
        k_value_raw = excel_handler.read_cell(15, 2)
        print(f"Read k_value: {k_value_raw} (type: {type(k_value_raw)})")  # Debugging print

        # Convert to float
        k_value = float(k_value_raw)
        print(f"Converted k_value: {k_value} (type: {type(k_value)})")  # Debugging print
    except Exception as e:
        print(f"Error reading or converting k_value: {e}")
        return None  # Stop initialization if there's an error

    # Send k_value to probe with error handling
    try:
        probe_device.send_set(k_value)
        print("k_value successfully sent to probe.")
    except Exception as e:
        print(f"Error while sending k_value to probe: {e}")
        return None  # Stop initialization if there's an error

    probe_device.send_show()
    probe_device.send_v()
    probe_device.send_binary(1)  # Enable binary output

    print("Probe initialized.") # Printed when succesful 
    return probe_device, k_value

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

    All the inputs are hardcoded based on the excel set up this is for program 1 user inputs. Column 7
    
    WARNING: Changing a cell in excel will raise an error and it will not allow to run the program via excel. 
    """
    
    print("Initializing experimental Parameters...")
    
    parameters = {
        "reading_interval": float(excel_handler.read_cell(3, 7)),
        "pre_run_voltage": float(excel_handler.read_cell(4, 7)),
        "pre_run_current": float(excel_handler.read_cell(5, 7)),
        "pre_run_time": float(excel_handler.read_cell(6, 7)),
        "pre_run_flowrate": float(excel_handler.read_cell(7, 7)),

        "segment1_voltage": float(excel_handler.read_cell(9, 7)),
        "segment1_current": float(excel_handler.read_cell(10, 7)),
        "segment1_time": float(excel_handler.read_cell(11, 7)),
        "segment1_flowrate": float(excel_handler.read_cell(12, 7)),

        "segment2_voltage": float(excel_handler.read_cell(19, 7)),
        "segment2_current": float(excel_handler.read_cell(20, 7)),
        "segment2_time": float(excel_handler.read_cell(21, 7)),
        "segment2_flowrate": float(excel_handler.read_cell(22, 7)),

        "post_run_voltage": float(excel_handler.read_cell(29, 7)),
        "post_run_current": float(excel_handler.read_cell(30, 7)),
        "post_run_time": float(excel_handler.read_cell(31, 7)),
        "post_run_flowrate": float(excel_handler.read_cell(32, 7)),

        "number_of_cycles": float(excel_handler.read_cell(29, 3)),
    }
            

    return parameters

def collect_and_log_data(excel_handler, power_supply, probe_device, overall_start_time, segment_name, flowrate, reading_interval, last_logged_time, cycle_count, k_value):
    """Collects and logs data from devices to the Excel file."""
    
    current_time = time.monotonic()
    elapsed_since_last_log = current_time - last_logged_time
    
    if elapsed_since_last_log >= reading_interval:
        elapsed_time = round(current_time - overall_start_time, 1)  # Rounds to 1 decimal place
        
        #Using method to aquire voltage and current readings

        voltage = power_supply.measure_voltage()
        current = power_supply.measure_current()
        
        #Converting conductivity based on correct conversion, fixing issue of incorrect conductivity 
        conductivity = probe_device.send_v()  
        probe_reading = round(conductivity * k_value * 1000, 3)
        
        # Write data to Excel sending to the correct method as the way the data in collected and written
        excel_handler.write_data_program1_2(elapsed_time, voltage, current, probe_reading, segment_name, flowrate, cycle_count)

        return current_time  # Update the last logged time to current time
    return last_logged_time  # No log, return the previous last logged time
  
def clear_all_unwanted(excel_handler):
    """
    Cleans up the Excel workbook by removing unused data, buttons, 
    and extra sheets to ensure only relevant information remains for the specific program being used. 
    
    Might need to modify buttons if they are label differently when downloaded. Macros is commented 
    out as it not relevant and the program can still run with out deleting specific macros.

    """
    print("Clearing cells and borders.")
    ranges_to_clear = [
        "I1:I38", "J1:J38",
        "L1:L38", "M1:M38",
        "O1:O38", "P1:P38",
        "Q1:Q38", "R1:R38",
        "S1:S38", "A18:A25",
        "B19:B25","C19:C25"
    ]
    for cell_range in ranges_to_clear:
        excel_handler.clear_cells("Inputs", cell_range)

    #print("Clearing macros")
   # excel_handler.delete_unused_macros(keep_modules=["Module7", "Module2", "Module9"])

    print("Clearing buttons")
    excel_handler.delete_unwanted_buttons("Inputs", keep_buttons=["Button 7", "Button 5", "Button 10"]) #Keeps only these buttons 

    print("Deleting sheets")
    sheets_to_delete = ["Program_2", "Program_3", "Program_4"]  #Deletes these sheets
    for sheet_name in sheets_to_delete:
        excel_handler.delete_sheet(sheet_name)
  
def main():
    """
    Runs the CDI experiment from start to finish for Program 1

    - Initializes the Excel workbook and clears unwanted data.
    - Connects to the power supply, pump, and probe.
    - Reads experiment parameters from the workbook.
    - Executes Pre-Run, repeated Segment cycles, and Post-Run.
    - Logs data at the configured interval and updates cycle counts.
    - Monitors a stop signal for safe termination.
    - Cleans up devices, threads, and files after completion or error.
    """
    close_open_terminals()  # Close other terminals at the start of the script

    # Check if the stop_signal.txt file exists and delete it (cleaning up from previous runs)
    if os.path.exists(STOP_SIGNAL_FILE):
        os.remove(STOP_SIGNAL_FILE)
        print(f"{STOP_SIGNAL_FILE} from a previous experiment has been removed.")

    # Check if workbook path is provided as an argument
    if len(sys.argv) != 2:
        print("Usage: python main.py <workbook_path>")
        sys.exit(1)

    workbook_path = sys.argv[1]

    
    excel_handler = None
    satellite_number = None

    try:
        # Initialize Excel Handler with provided workbook path 
        excel_handler = ExcelHandler(workbook_path, "Program_1")
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

        # Initialize all devices and experiment parameters
        power_supply, segment1_voltage, segment1_current = initialize_power_supply(excel_handler, power_supply_resource)
        probe_device, k_value = initialize_probe(excel_handler, probe_port)
        pump_device, satellite_number = initialize_pump(excel_handler, pump_port)
        experiment_parameters = read_experiment_parameters(excel_handler)

        # Await user input to start experiment
        print("Starting experiment...")

        # Perform Pre-Run (Only once)
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
            last_logged_time = collect_and_log_data(excel_handler, power_supply, probe_device, overall_start_time, "Pre-Run", pre_run_flowrate, experiment_parameters["reading_interval"], last_logged_time, 0, k_value)

        # Create a threading event to signal stopping
        stop_event = threading.Event()

        # Start a thread to monitor the stop signal file
        file_thread = threading.Thread(target=monitor_file, args=(stop_event,))
        file_thread.start()
        
        # Retrieve the number of cycles , turning into integer to use in the main loop
        num_cycles = int(experiment_parameters["number_of_cycles"])  
        # Set cycle count to 0 this will change when the experiment starts
        cycle_count = 0

        # Main experimental phases (loop)
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
    
        stop_requested = False  # Track if stop was clicked
        cycle_count = 0  # starting cycle count
        
        # Main loop: 
        # runs segments repeatedly until stop signal is detected or cycles are complete
        while True:
           #prints out so users can see through the terminal
            print(f"Cycle Count: {cycle_count}, Stop Requested: {stop_requested}, Num Cycles: {num_cycles}")
        
            # Check if the stop signal is set or if all cycles have been completed
            if stop_requested or cycle_count >= num_cycles:
                print("Experiment finished or stop requested.")
                break  # Break the main loop when stop is requested or all cycles are completed
                
            #Prints which cycle is starting
            print(f"Starting Cycle {cycle_count + 1} of {num_cycles}...")
        
            # Loop through the segments as a whole (seg1, inter1, seg2, inter2)
            for segment, voltage, current, flowrate, duration in segments:
                print(f"Starting {segment}...")
                segment_start_time = time.perf_counter()
                power_supply.set_voltage_current(voltage, current)
                pump_device.convert_and_set_speed(satellite_number, flowrate)
                
                # Log data during the segment time
                if duration > 0:  # Don't log data for segments that don't have a set time, skip those log times 
                    while time.monotonic() - segment_start_time < duration:
                        last_logged_time = collect_and_log_data(
                            excel_handler, power_supply, probe_device, overall_start_time, 
                            segment, flowrate, experiment_parameters["reading_interval"], 
                            last_logged_time, cycle_count + 1, k_value
                        )
        
                        #Check if stop was requested through out the segments 
                        if stop_event.is_set():
                            stop_requested = True  # Mark that stop was requested to make sure to end cycle once intersegment 2 is completed

                if stop_requested and segment == "Segment 2":
                    break  # Exit the main loop after completing Inter-Segment 2
   
            # Update cycle count after each cycle is complete, regardless is stop request is set, it will still write what cycle it ended in
            cycle_count += 1
            excel_handler.write_cycle_count(cycle_count) #Update cycle count after each cycle is completed, live feedback for users
            print(f"Cycle {cycle_count} completed.")
        
        # Perform Post-Run (Only once)
        print("Starting Post-Run...")
        post_run_voltage = experiment_parameters["post_run_voltage"]
        post_run_current = experiment_parameters["post_run_current"]
        post_run_flowrate = experiment_parameters["post_run_flowrate"]
        post_run_time = experiment_parameters["post_run_time"]
        power_supply.set_voltage_current(post_run_voltage, post_run_current)
        pump_device.convert_and_set_speed(satellite_number, post_run_flowrate)
        post_run_start_time = time.monotonic()
        
        #Collecting post run data
        while time.monotonic() - post_run_start_time < post_run_time:
            last_logged_time = collect_and_log_data(excel_handler, power_supply, probe_device, overall_start_time, "Post-Run", post_run_flowrate, experiment_parameters["reading_interval"], last_logged_time, 0, k_value)
        
        print("Experiment completed.")
        excel_handler.save_wb()  # Save data to the workbook 

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        stop_experiment(pump_device, power_supply, probe_device, excel_handler, satellite_number)
        stop_event.set()
        file_thread.join()  # Wait for the file monitoring thread to finish
        if os.path.exists(STOP_SIGNAL_FILE):
            os.remove(STOP_SIGNAL_FILE)  # Clean up stop signal file if it exists

if __name__ == '__main__':
    main()





