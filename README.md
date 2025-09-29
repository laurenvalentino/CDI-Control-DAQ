# CDI Python-Based Program

**Author:** Molly S. Perez  
**Mentor:** Dr. Lauren Valentino

## Table of Contents
1. [Overview](#overview)  
2. [Features](#features)  
3. [Setup & Usage](#Setup&Usage)
4. [File Structure](#file-structure)
5. [Contribution](#contribution)
6. [License](#license)  

---

## Overview
The **CDI Program** is a laboratory automation tool designed to control and monitor electrochemical systems. The tool was developed for capacitive deionization (CDI)
experiments, but it can be used for any system that requires controlled voltage or current segments and multi-parameter monitoring. The program integrates hardware components
to run user-defined experimental parameters, providing operational control of a programmable power supply, peristaltic pump, and data acquisition devices.

---

## Features
1. Automatic hardware control of power supply, peristaltic pump, and probes 
2. Multiple experimental workflows with 2 or 4 segments per cycle
3. Real-time monitoring and plotting of voltage, current, flow rate, conductivity, pH, and temperature
4.  Automated cycle control with initialization and post-run stabilization
5.  Data logging and export to a single synchronized Excel file
6.   Dual-cell control for simultaneous experiments under different voltage/current conditions
7.   User-friendly Excel-based interface for modifying experimental inputs without editing code  


---

## Setup & Usage
For detailed instructions on installation, program usage, dependencies, and configuration, please refer to the **CDI Program Manual.pdf**.

---

## File Structure
    -CDI_Programs_/
      -CDI Program_.xlsm
      -CDI_Python_Program_Files/
        -ExcelHandler.py
        -multi_edaq_probe_working.py
        -probe_working.py
        -psu.py
        -pump_initialize.py
        -Program_1_Main.py
        -Program_2_Main.py
        -Program_3_Main.py
        -Program_4_Main.py
    -CDI Program Manual.pdf
  
    
---

## Contribution
The development of this program was carried out by [Molly Perez](https://github.com/perezmolly973) with supervision and guidance from Dr. Lauren Valentino.

---

## License
Left this blank to add this information once the copyright application is approved

---

