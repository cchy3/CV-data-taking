#!/usr/local/bin/python3
# -*- coding: utf-8 -*-
import sys
from Keithley import Keithley2400, Keithley2657a
from Agilent import AgilentE4980a, Agilent4156
from emailbot import sendMail
from tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu
from tkinter import ttk
from tkinter.constants import LEFT, RIGHT, RAISED
import matplotlib
import threading
from random import randint
from _ast import Param
from platform import platform
matplotlib.use("TkAgg")
import platform

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import pyplot as plt
import time, math
import visa
import tkinter.filedialog
import xlsxwriter
import queue as Queue
import random
import json
import numpy as np


debug = False
SAVE_FILE = ".labmater_v5_settings.json"
rm = visa.ResourceManager()
print(rm.list_resources())

cv_start_time = None
global_end_volt = None
sensor_params = [
    "temperature",
    "humidity",
    "your_name",
    "vendor",
    "product_name",
    "Type",
    "geometry",
    "wafer",
    "sensor_type",
    "chip_num",
    "fluence",
    "fluence_type",
    "area",
    "set_number"
]


def getCVSettings(gui):

    cv = {}
    cv["compliance"] = gui.cv_compliance.get()
    cv["complianceScale"] = gui.cv_compliance_scale.get()
    
    cv["startVoltage1"] = gui.cv_start_volt1.get()
    cv["stepVoltage1"] = gui.cv_step_volt1.get()
    cv["holdTime1"] = gui.cv_hold_time1.get()

    cv["startVoltage2"] = gui.cv_start_volt2.get()
    cv["stepVoltage2"] = gui.cv_step_volt2.get()
    cv["holdTime2"] = gui.cv_hold_time2.get()

    cv["startVoltage3"] = gui.cv_start_volt3.get()
    cv["stepVoltage3"] = gui.cv_step_volt3.get()
    cv["holdTime3"] = gui.cv_hold_time3.get()
    
    cv["endVoltage"] = gui.cv_end_volt.get()
    
    cv["device"] = gui.cv_source_choice.get()
    cv["freq"] = gui.cv_frequencies.get().split(",")
    cv["function"] = gui.cv_function_choice.get()
    cv["amplitude"] = gui.cv_amplitude.get()
    cv["impedance"] = gui.cv_impedance.get()
    cv["integration"] = gui.cv_integration.get()
    cv["instability"] = gui.cv_instability.get()
    cv["instability_wait_time"] = gui.cv_instability_wait_time.get()
    cv["recipients"] = gui.cv_recipients.get()

    global global_end_volt
    global_end_volt = -1*float(gui.cv_end_volt.get())

    global sensor_params
    for item in sensor_params:
        # print("get CV {}".format(item))
        cv[item] = getattr(gui, "cv_{}".format(item)).get()
        # print(" get CV {}".format(cv[item]))

    return cv


def setCVSettings(gui, cv):
    # print("-----------running setCVSettings")
    gui.cv_compliance.set(cv["compliance"])
    gui.cv_compliance_scale.set(cv["complianceScale"])

    gui.cv_start_volt1.set(cv["startVoltage1"])
    gui.cv_step_volt1.set(cv["stepVoltage1"])
    gui.cv_hold_time1.set(cv["holdTime1"])
    
    gui.cv_start_volt2.set(cv["startVoltage2"])
    gui.cv_step_volt2.set(cv["stepVoltage2"])
    gui.cv_hold_time2.set(cv["holdTime2"])
    
    gui.cv_start_volt3.set(cv["startVoltage3"])
    gui.cv_step_volt3.set(cv["stepVoltage3"])
    gui.cv_hold_time3.set(cv["holdTime3"])
    gui.cv_end_volt.set(cv["endVoltage"])
    
    gui.cv_source_choice.set(cv["device"])
    gui.cv_function_choice.set(cv["function"])
    gui.cv_amplitude.set(cv["amplitude"])
    gui.cv_impedance.set(cv["impedance"])
    gui.cv_integration.set(cv["integration"])
    gui.cv_recipients.set(cv["recipients"])
    gui.cv_instability.set(cv["instability"])
    gui.cv_instability_wait_time.set(cv["instability_wait_time"])
    # print("running setCVSettings loop")
    global sensor_params
    for item in sensor_params:
        # print("setCV {}".format(item))
        # print(cv[item])
        getattr(gui, "cv_{}".format(item)).set(cv[item])
        # print("end setCV {}".format(item))

    freq = ""
    for f in cv["freq"]:
        freq += f + ","
    freq = freq[:-1]
    gui.cv_frequencies.set(freq)

    return


def saveSettings(gui):
    settings = {"cv": getCVSettings(gui), "iv": None}
    # print("-----------------print in saveSetting")
    # print(settings["cv"])
    # print("-----------------print in saveSetting")
    try:
        # print("====>saveSetting")
        with open(SAVE_FILE, "w+") as f:
            f.write(json.dumps(settings))
            print("Settings saved in the file %s" % SAVE_FILE)
    except Exception as e:
        print("catch exception in saveSetting")
        print(e)
    print("end of saveSetting")
    return


def loadSettings(gui):
    settings = None

    print("Loading save.")
    with open(SAVE_FILE, "r") as f:
        f.seek(0)
        settings = json.loads(f.read())
        # print(settings['cv'])
        # print("here:")
    try:
        setCVSettings(gui, settings["cv"])
    except Exception as e:
        print("catch exception in loadSettings")
        print(e)


def GetIV(sourceparam, sourcemeter, dataout, stopqueue):
    (start_volt, end_volt, step_volt, delay_time, compliance) = sourceparam

    currents = []
    voltages = []
    keithley = 0

    if debug:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            keithley = Keithley2657a()
        keithley.configure_measurement(1, 0, compliance)
    last_volt = 0
    badCount = 0

    scaled = False
    if step_volt < 1.0:
        start_volt *= 1000
        end_volt *= 1000
        step_volt *= 1000
        scaled = True

    if start_volt > end_volt:
        step_volt = -1 * step_volt

    print("looping now")

    for volt in np.arange(start_volt, end_volt, int(step_volt)):
        if not stopqueue.empty():
            stopqueue.get()
            break
        start_time = time.time()

        curr = 0
        if debug:
            pass
        else:
            if scaled:
                keithley.set_output(volt / 1000.0)
            else:
                keithley.set_output(volt)

        time.sleep(delay_time)

        if debug:
            curr = (volt + randint(0, 10)) * 1e-9
        else:
            curr = keithley.get_current()

        print("Current Reading: " + str(curr))

        # curr = volt
        time.sleep(1)
        if abs(curr) > abs(compliance - 50e-9):
            badCount = badCount + 1
        else:
            badCount = 0

        if badCount >= 5:
            print("Compliance reached")
            dataout.put(((voltages, currents), 100, 0))
            break

        currents.append(curr)
        if scaled:
            voltages.append(volt / 1000.0)
        else:
            voltages.append(volt)

        if scaled:
            last_volt = volt / 1000.0
        else:
            last_volt = volt

        time_remain = (time.time() - start_time) * (abs((end_volt - volt) / step_volt))

        dataout.put(
            (
                (voltages, currents),
                100 * abs((volt + step_volt) / float(end_volt)),
                time_remain,
            )
        )

    while abs(last_volt) > 25:
        if debug:
            pass
        else:
            keithley.set_output(last_volt)

        time.sleep(delay_time / 2.0)

        if last_volt < 0:
            last_volt += abs(step_volt * 2.0)
        else:
            last_volt -= abs(step_volt * 2.0)

    time.sleep(delay_time / 2.0)
    if debug:
        pass
    else:
        keithley.set_output(0)
        keithley.enable_output(False)
    return (voltages, currents)

def cv_DifferentVoltageRange(dataout, scaled, level, compliance, start_time, startV, endV, stepV, holdTime, frequencies, voltages, capacitance, c, p2,
                             keithley, agilent, instability, instability_wait_time, badCount, stopqueue):
    capacitance = []
    voltages = []
    p2 = []
    c = []
    for volt in np.arange(startV, endV+int(stepV), int(stepV)):
        if not stopqueue.empty():
            print("here1")
            stopqueue.get()
            break

        start_time = time.time()
        if debug:
            pass
        else:
            if scaled:
                keithley.set_output(volt / 1000.0)
            else:
                keithley.set_output(volt)

        curr = 0
        for f in frequencies:
            time.sleep(holdTime)

            if debug:
                capacitance.append((volt + int(f) * randint(0, 10)))
                curr = volt * 1e-10
                c.append(curr)
                p2.append(volt * 10)
            else:
                agilent.configure_measurement_signal(float(f), 0, level)
                if instability == 0:
                    (data, aux) = agilent.read_data()
                    capacitance.append(data)
                else:
                    print(
                        "Checking the instability of the measurment. Threshold: {}%".format(
                            100 * instability
                        )
                    )
                    pastCapacitance = None
                    currentCapacitance = None
                    percentage = 100
                    notReady = pastCapacitance is None or currentCapacitance is None
                    while notReady or percentage > instability:
                        if not stopqueue.empty():
                            stopqueue.get()
                            break
                        pastCapacitance = currentCapacitance
                        (currentCapacitance, aux) = agilent.read_data()
                        notReady = pastCapacitance is None or currentCapacitance is None
                        if not notReady:
                            percentage = abs(
                                (currentCapacitance - pastCapacitance) / pastCapacitance
                            )
                            print(
                                "Instability {:8.02f}%; pause for {:8.02f} sec.".format(
                                    100 * percentage, instability_wait_time
                                )
                            )
                            time.sleep(instability_wait_time)  # Heyi Delay
                    capacitance.append(currentCapacitance)

                p2.append(aux)
                curr = keithley.get_current()
                c.append(curr)

        if abs(curr) > abs(compliance - 50e-9):
            badCount = badCount + 1
        else:
            badCount = 0

        if badCount >= 5:
            print("Compliance reached")
            break

        time_remain = (time.time() - start_time) * (abs((endV - volt) / stepV))

        if scaled:
            voltages.append(volt / 1000.0)
        else:
            voltages.append(volt)

        formatted_cap = []
        parameter2 = []
        currents = []
        
        global global_end_volt
        
        for i in np.arange(0, len(frequencies), 1):
            formatted_cap.append(capacitance[i :: len(frequencies)])
            parameter2.append(p2[i :: len(frequencies)])
            currents.append(c[i :: len(frequencies)])
        dataout.put(
            (
                (voltages, formatted_cap),
                100 * abs((volt + stepV) / (global_end_volt*1000)),
                time_remain,
            )
        )

        time_remain = time.time() + (time.time() - start_time) * (
            abs((volt - endV) / endV)
        )

        if scaled:
            last_volt = volt / 1000.0
        else:
            last_volt = volt
    return voltages, currents, formatted_cap, parameter2

def GetCV(params, sourcemeter, dataout, stopqueue):

    capacitance = []
    voltages = []
    p2 = []
    c = []
    keithley = 0
    agilent = 0
    if debug:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            print("Creating a 2657a object.")
            keithley = Keithley2657a()

    last_volt = 0

    (
        start_volt1,
        step_volt1,
        delay_time1,
        start_volt2,
        step_volt2,
        delay_time2,
        start_volt3,
        step_volt3,
        delay_time3,
        end_volt,
        
        compliance,
        frequencies,
        level,
        function,
        impedance,
        int_time,
        instability,
        instability_wait_time,
        temperature,
        humidity,
        your_name,
        vendor,
        product_name,
        Type,
        geometry,
        wafer,
        sensor_type,
        chip_num,
        fluence,
        fluence_type,
        area,
        set_number
    ) = params
    if debug:
        pass
    else:
        keithley.configure_measurement(1, 0, compliance)

    if debug:
        pass
    else:
        agilent = AgilentE4980a()
        agilent.configure_measurement(function)
        agilent.configure_aperture(int_time)
    badCount = 0  

    scaled = False

    if step_volt1 < 1.0: #in params, start volts and end volt are int, so this is to make sure it starts from 0 or not?
        start_volt1 *= 1000
        step_volt1 *= 1000
        start_volt2 *= 1000
        step_volt2 *= 1000
        start_volt3 *= 1000
        step_volt3 *= 1000       
        end_volt *= 1000
        
        scaled = True

    if start_volt1 > end_volt:
        if start_volt2 > end_volt:
            if start_volt3 > end_volt:
                step_volt1 = -1 * step_volt1
                step_volt2 = -1 * step_volt2
                step_volt3 = -1 * step_volt3
            else:
                print("The end volt must be negative, and the magnitude should be bigger than the magnitude of the start volt 3!")

        else:
            print("The end volt must be negative, and the magnitude should be bigger than the magnitude of the start volt 2!")

    else:
        print("The end volt must be negative, and the magnitude should be bigger than the magnitude of the start volt 1!")

        
    start_time = time.time()
    
    #---------------------------for loop starts here
    _voltages =[]
    _currents =[]
    _formatted_cap=[]
    _parameter2 =[]
    
    temp_data_out = cv_DifferentVoltageRange(dataout, scaled, level, compliance, start_time, start_volt1, start_volt2, step_volt1, delay_time1, frequencies, voltages, capacitance, c, p2, keithley, agilent, instability, instability_wait_time, badCount, stopqueue)
    _voltages.append(temp_data_out[0])
    _currents.append(temp_data_out[1])
    _formatted_cap.append(temp_data_out[2])
    _parameter2.append(temp_data_out[3])

    
    temp_data_out = cv_DifferentVoltageRange(dataout, scaled, level, compliance, start_time, start_volt2, start_volt3, step_volt2, delay_time2, frequencies, voltages, capacitance, c, p2, keithley, agilent, instability, instability_wait_time, badCount, stopqueue)
    _voltages.append(temp_data_out[0])
    _currents.append(temp_data_out[1])
    _formatted_cap.append(temp_data_out[2])
    _parameter2.append(temp_data_out[3])


    temp_data_out = cv_DifferentVoltageRange(dataout, scaled, level, compliance, start_time, start_volt3, end_volt, step_volt3, delay_time3, frequencies, voltages, capacitance, c, p2, keithley, agilent, instability, instability_wait_time, badCount, stopqueue)
    _voltages.append(temp_data_out[0])
    _currents.append(temp_data_out[1])
    _formatted_cap.append(temp_data_out[2])
    _parameter2.append(temp_data_out[3])

    
    #----------------------------for loop end
        
    if scaled:
        last_volt = last_volt / 1000

    if debug:
        pass
    else:
        keithley.powerDownPSU()
        keithley.enable_output(False)

    return (_voltages, _currents, _formatted_cap, _parameter2)


def spa_iv(params, dataout, stopqueue):
    (start_volt, end_volt, step_volt, hold_time, compliance, int_time) = params

    print(params)
    voltage_smua = []
    current_smua = []

    current_smu1 = []
    current_smu2 = []
    current_smu3 = []
    current_smu4 = []
    voltage_vmu1 = []

    voltage_source = Keithley2657a()
    voltage_source.configure_measurement(1, 0, compliance)
    voltage_source.enable_output(True)

    daq = Agilent4156()
    daq.configure_integration_time(_int_time=int_time)

    scaled = False
    if step_volt < 1.0:
        start_volt *= 1000
        end_volt *= 1000
        step_volt *= 1000
        scaled = True

    if start_volt > end_volt:
        step_volt = -1 * step_volt

    for i in np.arange(0, 4, 1):
        daq.configure_channel(i)
    daq.configure_vmu()

    last_volt = 0
    for volt in np.arange(start_volt, end_volt, step_volt):

        if debug:
            pass
        else:
            if scaled:
                voltage_source.set_output(volt / 1000.0)
            else:
                voltage_source.set_output(volt)
        time.sleep(hold_time)

        daq.configure_measurement()
        daq.configure_sampling_measurement()
        daq.configure_sampling_stop()

        # daq.inst.write(":PAGE:DISP:GRAP:Y2:NAME \'I2\';")
        daq.inst.write(":PAGE:DISP:LIST '@TIME', 'I1', 'I2', 'I3', 'I4', 'VMU1'")
        daq.measurement_actions()
        daq.wait_for_acquisition()

        current_smu1.append(daq.read_trace_data("I1"))
        current_smu2.append(daq.read_trace_data("I2"))

        # daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I2\', \'I3\'")

        current_smu3.append(daq.read_trace_data("I3"))
        current_smu4.append(daq.read_trace_data("I4"))
        voltage_vmu1.append(daq.read_trace_data("VMU1"))
        current_smua.append(voltage_source.get_current())

        if scaled:
            voltage_smua.append(volt / 1000.0)
            last_volt = volt / 1000.0
        else:
            voltage_smua.append(volt)
            last_volt = volt

        print("SMU1-4")
        print(current_smu1)
        print(current_smu2)
        print(current_smu3)
        print(current_smu4)
        print("SMUA")
        print(current_smua)
        print("VMU1")
        print(voltage_vmu1)
        dataout.put(
            (
                voltage_vmu1,
                current_smua,
                current_smu1,
                current_smu2,
                current_smu3,
                current_smu4,
            )
        )
    while abs(last_volt) >= 4:
        time.sleep(0.5)

        if debug:
            pass
        else:
            voltage_source.set_output(last_volt)

        if last_volt < 0:
            last_volt += 5
        else:
            last_volt -= 5

    time.sleep(0.5)
    voltage_source.set_output(0)
    voltage_source.enable_output(False)
    return (
        voltage_smua,
        current_smua,
        current_smu1,
        current_smu2,
        current_smu3,
        current_smu4,
        voltage_vmu1,
    )


# TODO: current monitor bugfixes and fifo implementation
def curmon(source_params, sourcemeter, dataout, stopqueue):

    (voltage_point, step_volt, hold_time, compliance, minutes) = source_params
    print("(voltage_point, step_volt, hold_time, compliance, minutes)")
    print(source_params)
    currents = []
    timestamps = []
    voltages = []

    total_time = minutes * 60

    keithley = 0
    if debug:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            keithley = Keithley2657a()
        keithley.configure_measurement(1, 0, compliance)

    last_volt = 0
    badCount = 0

    scaled = False

    if step_volt < 1:
        voltage_point *= 1000
        step_volt *= 1000
        scaled = True
    else:
        step_volt = int(step_volt)

    if 0 > voltage_point:
        step_volt = -1 * step_volt

    start_time = time.time()

    for volt in np.arange(0, voltage_point, step_volt):
        if not stopqueue.empty():
            stopqueue.get()
            break

        curr = 0
        if debug:
            pass
        else:
            if scaled:
                keithley.set_output(volt / 1000.0)
            else:
                keithley.set_output(volt)

        time.sleep(hold_time)

        if debug:
            curr = (volt + randint(0, 10)) * 1e-9
        else:
            curr = keithley.get_current()
        # curr = volt

        if abs(curr) > abs(compliance - 50e-9):
            badCount = badCount + 1
        else:
            badCount = 0

        if badCount >= 5:
            print("Compliance reached")
            break

        if scaled:
            last_volt = volt / 1000.0
        else:
            last_volt = volt

        dataout.put(((timestamps, currents), 0, total_time + start_time))
        print("""ramping up""")

    print("current time")
    print(time.time())
    print("Start time")
    print(start_time)
    print("total time")
    print(total_time)

    start_time = time.time()
    while time.time() < start_time + total_time:
        time.sleep(5)

        dataout.put(
            (
                (timestamps, currents),
                100 * ((time.time() - start_time) / total_time),
                start_time + total_time,
            )
        )
        if debug:
            currents.append(randint(0, 10) * 1e-9)
        else:
            currents.append(keithley.get_current())
        timestamps.append(time.time() - start_time)
        print("timestamps")
        print(timestamps)
        print("currents")
        print(currents)
    print("Finished")

    while abs(last_volt) > 5:
        if debug:
            pass
        else:
            keithley.set_output(last_volt)

        time.sleep(hold_time / 2.0)
        if last_volt < 0:
            last_volt += 5
        else:
            last_volt -= 5

    time.sleep(hold_time / 2.0)
    if debug:
        pass
    else:
        keithley.set_output(0)
        keithley.enable_output(False)

    return (timestamps, currents)



rownumber = 0
colnumber = 1


class GuiPart():
    def AddEntry(self, sw, in_item, same_row, label, default_text):
        global rownumber
        global colnumber
        getattr(self, "cv_{}".format(in_item)).set(default_text)
        sw = Label(self.f2, text=label)
        sw.grid(row=rownumber, column=colnumber)
        colnumber += 1
        sw = Entry(self.f2, textvariable=getattr(self, "cv_{}".format(in_item)))
        sw.grid(row=rownumber, column=colnumber)
        colnumber += 1
        if same_row == False:
            rownumber += 1
            colnumber = 1
        # print(rownumber, colnumber)

    def AddButton(self, sw, in_item, same_row, label, choices, default_choice):

        global rownumber
        global colnumber
        setattr(self, "cv_{}".format(in_item), StringVar())
        getattr(self, "cv_{}".format(in_item)).set(default_choice)
        sw = Label(self.f2, text=label)
        sw.grid(row=rownumber, column=colnumber)
        colnumber += 1
        sw = OptionMenu(self.f2, getattr(self, "cv_{}".format(in_item)), *choices)
        sw.grid(row=rownumber, column=colnumber)

        colnumber += 1
        if same_row == False:
            rownumber += 1
            colnumber = 1
        # print(rownumber, colnumber)

    def __init__(self, master, inputdata, outputdata, stopq):
        print("in guipart")

        self.master = master
        self.inputdata = inputdata
        self.outputdata = outputdata
        self.stop = stopq

        self.start_volt = StringVar()
        self.end_volt = StringVar()
        self.step_volt = StringVar()
        self.hold_time = StringVar()
        
        self.compliance = StringVar()
        self.recipients = StringVar()
        self.compliance_scale = StringVar()
        self.source_choice = StringVar()
        self.filename = StringVar()

        self.cv_filename = StringVar()
        
        self.cv_start_volt1 = StringVar()
        self.cv_step_volt1 = StringVar()
        self.cv_hold_time1 = StringVar()
        self.cv_start_volt2 = StringVar()
        self.cv_step_volt2 = StringVar()
        self.cv_hold_time2 = StringVar()
        self.cv_start_volt3 = StringVar()
        self.cv_step_volt3 = StringVar()
        self.cv_hold_time3 = StringVar()
        self.cv_end_volt = StringVar()
        
        self.cv_compliance = StringVar()
        self.cv_instability = StringVar()
        self.cv_instability_wait_time = StringVar()

        global sensor_params
        for item in sensor_params:
            if item !="fluence_type" or item != "vendor" or item != "geometry":
                setattr(self, "cv_{}".format(item), StringVar())

        self.cv_recipients = StringVar()
        self.cv_compliance_scale = StringVar()
        self.cv_source_choice = StringVar()
        self.cv_impedance_scale = StringVar()
        self.cv_amplitude = StringVar()
        self.cv_frequencies = StringVar()
        self.cv_integration = StringVar()
        self.started = False

        self.multiv_start_volt = StringVar()
        self.multiv_end_volt = StringVar()
        self.multiv_step_volt = StringVar()
        self.multiv_hold_time = StringVar()
        self.multiv_compliance = StringVar()
        self.multiv_recipients = StringVar()
        self.multiv_compliance_scale = StringVar()
        self.multiv_source_choice = StringVar()
        self.multiv_filename = StringVar()
        self.multiv_times = StringVar()

        self.curmon_start_volt = StringVar()
        self.curmon_end_volt = StringVar()
        self.curmon_step_volt = StringVar()
        self.curmon_hold_time = StringVar()
        self.curmon_compliance = StringVar()
        self.curmon_recipients = StringVar()
        self.curmon_compliance_scale = StringVar()
        self.curmon_source_choice = StringVar()
        self.curmon_filename = StringVar()
        self.curmon_time = StringVar()

        """
        IV GUI
        """
        ""
        self.start_volt.set("0.0")
        self.end_volt.set("100.0")
        self.step_volt.set("5.0")
        self.hold_time.set("1.0")
        self.compliance.set("1.0")

        self.f = plt.figure(figsize=(6, 4), dpi=60)
        self.a = self.f.add_subplot(111)

        self.cv_f = plt.figure(figsize=(6, 4), dpi=60)
        self.cv_a = self.cv_f.add_subplot(111)

        self.multiv_f = plt.figure(figsize=(6, 4), dpi=60)
        self.multiv_a = self.multiv_f.add_subplot(111)

        self.curmon_f = plt.figure(figsize=(6, 4), dpi=60)
        self.curmon_a = self.curmon_f.add_subplot(111)

        n = ttk.Notebook(root, width=800)
        n.grid(row=0, column=0, columnspan=100, rowspan=100, sticky="NESW")
        self.f1 = ttk.Frame(n)
        self.f2 = ttk.Frame(n)
        self.f3 = ttk.Frame(n)
        self.f4 = ttk.Frame(n)
        self.f5 = ttk.Frame(n)
        n.add(self.f1, text="Basic IV")
        n.add(self.f2, text="CV")
        n.add(self.f3, text="Param Analyzer IV ")
        n.add(self.f4, text="Multiple IV")
        n.add(self.f5, text="Current Monitor")

        if "Windows" in platform.platform():
            self.filename.set("LabMasterData\iv_data")
            s = Label(self.f1, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(self.f1, textvariable=self.filename)
            s.grid(row=0, column=2)

        s = Label(self.f1, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f1, textvariable=self.start_volt)
        s.grid(row=1, column=2)
        s = Label(self.f1, text="V")
        s.grid(row=1, column=3)

        s = Label(self.f1, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f1, textvariable=self.end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f1, text="V")
        s.grid(row=2, column=3)

        s = Label(self.f1, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f1, textvariable=self.step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f1, text="V")
        s.grid(row=3, column=3)

        s = Label(self.f1, text="Hold Time")
        s.grid(row=4, column=1)
        s = Entry(self.f1, textvariable=self.hold_time)
        s.grid(row=4, column=2)
        s = Label(self.f1, text="s")
        s.grid(row=4, column=3)

        s = Label(self.f1, text="Compliance")
        s.grid(row=5, column=1)
        s = Entry(self.f1, textvariable=self.compliance)
        s.grid(row=5, column=2)
        compliance_choices = {"mA", "uA", "nA"}
        self.compliance_scale.set("uA")
        s = OptionMenu(self.f1, self.compliance_scale, *compliance_choices)
        s.grid(row=5, column=3)

        self.recipients.set("adapbot@gmail.com")
        s = Label(self.f1, text="Email data to:")
        s.grid(row=6, column=1)
        s = Entry(self.f1, textvariable=self.recipients)
        s.grid(row=6, column=2)

        source_choices = {"Keithley 2400", "Keithley 2657a"}
        self.source_choice.set("Keithley 2657a")
        s = OptionMenu(self.f1, self.source_choice, *source_choices)
        s.grid(row=0, column=7)

        s = Label(self.f1, text="Progress:")
        s.grid(row=11, column=1)

        s = Label(self.f1, text="Est finish at:")
        s.grid(row=12, column=1)

        timetext = str(time.asctime(time.localtime(time.time())))
        self.timer = Label(self.f1, text=timetext)
        self.timer.grid(row=12, column=2)

        self.pb = ttk.Progressbar(
            self.f1, orient="horizontal", length=200, mode="determinate"
        )
        self.pb.grid(row=11, column=2, columnspan=5)
        self.pb["maximum"] = 100
        self.pb["value"] = 0

        self.canvas = FigureCanvasTkAgg(self.f, master=self.f1)
        self.canvas.get_tk_widget().grid(row=7, columnspan=10)
        self.a.set_title("IV")
        self.a.set_xlabel("Voltage")
        self.a.set_ylabel("Current")

        # plt.xlabel("Voltage")
        # plt.ylabel("Current")
        # plt.title("IV")
        self.canvas.draw()

        s = Button(self.f1, text="Start IV", command=self.prepare_values)
        s.grid(row=3, column=7)

        s = Button(self.f1, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        """
        /***********************************************************
         * CV GUI
         **********************************************************/
        """
        print("setting CV GUI")
        self.cv_filename = StringVar()
        self.cv_filename.set("LabMasterData\cv_data")
        global rownumber
        global colnumber

        if "Windows" in platform.platform():
            s = Label(self.f2, text="File name")
            s.grid(row=rownumber, column=1)
            s = Entry(self.f2, textvariable=self.cv_filename)
            s.grid(row=rownumber, column=2)

        rownumber += 1
        self.AddEntry(s, sensor_params[2], False, "Your Name", "Enter your name here")
        self.AddEntry(s, sensor_params[0], True, "Temperature", "25.0")
        self.AddEntry(s, sensor_params[1], False, "Humidity", "40.0")
        self.AddButton(
            s, sensor_params[3], True, "Vendor", {"HPK", "FBK", "CNM"}, "HPK"
        )
        self.AddEntry(s, sensor_params[4], True, "Product Name", "i.e. HGTD")
        self.AddEntry(s, sensor_params[5], False, "Type", "i.e. 3.1")
        self.AddButton(
            s,
            sensor_params[6],
            True,
            "Geometry",
            {"single_pad", "2x2", "3x3", "5x5"},
            "single_pad",
        )
        self.AddEntry(s, sensor_params[7], True, "Wafer(W?)", "")
        self.AddEntry(s, sensor_params[13], False, "SET Number(P?)", "")
        self.AddEntry(s, sensor_params[8], True, "Sensor Type(SE?)", "")
        self.AddEntry(s, sensor_params[9], False, "Chip Number(#?)", "")
        self.AddEntry(s, sensor_params[10], True, "Fluence", "format: prerad, 1.5e16")
        self.AddButton(
            s,
            sensor_params[11],
            False,
            "Fluence Type",
            {"Non-irradiated", "JSI_neutron", "Cyric_proton", "LA_proton", "CERN_proton"},
            "JSI_neutron",
        )
        self.AddEntry(s, sensor_params[12], False, "Area", "?mm * ?mm")

        rownumber += 1

        self.cv_start_volt1.set("0.0")
        s = Label(self.f2, text="Start Volt1 [V]")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_start_volt1)
        s.grid(row=rownumber, column=2)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=3)
        self.cv_start_volt2.set("-50.0")
        s = Label(self.f2, text="Start Volt2 [V]")
        s.grid(row=rownumber, column=3)
        s = Entry(self.f2, textvariable=self.cv_start_volt2)
        s.grid(row=rownumber, column=4)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=6)
        self.cv_start_volt3.set("-70.0")
        s = Label(self.f2, text="Start Volt2 [V]")
        s.grid(row=rownumber, column=5)
        s = Entry(self.f2, textvariable=self.cv_start_volt3)
        s.grid(row=rownumber, column=6)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=9)
        rownumber += 1

        self.cv_step_volt1.set("1.0")
        s = Label(self.f2, text="Step Volt1 [V]")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_step_volt1)
        s.grid(row=rownumber, column=2)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=3)
        self.cv_step_volt2.set("1.0")
        s = Label(self.f2, text="Step Volt2 [V]")
        s.grid(row=rownumber, column=3)
        s = Entry(self.f2, textvariable=self.cv_step_volt2)
        s.grid(row=rownumber, column=4)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=6)
        self.cv_step_volt3.set("1.0")
        s = Label(self.f2, text="Step Volt3 [V]")
        s.grid(row=rownumber, column=5)
        s = Entry(self.f2, textvariable=self.cv_step_volt3)
        s.grid(row=rownumber, column=6)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=9)
        rownumber += 1

        self.cv_hold_time1.set("1.0")
        s = Label(self.f2, text="Hold Time1")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_hold_time1)
        s.grid(row=rownumber, column=2)
        #s = Label(self.f2, text="s")
        #s.grid(row=rownumber, column=3)
        self.cv_hold_time2.set("1.0")
        s = Label(self.f2, text="Hold Time2")
        s.grid(row=rownumber, column=3)
        s = Entry(self.f2, textvariable=self.cv_hold_time2)
        s.grid(row=rownumber, column=4)
        #s = Label(self.f2, text="s")
        #s.grid(row=rownumber, column=3)
        self.cv_hold_time3.set("1.0")
        s = Label(self.f2, text="Hold Time3")
        s.grid(row=rownumber, column=5)
        s = Entry(self.f2, textvariable=self.cv_hold_time3)
        s.grid(row=rownumber, column=6)
        #s = Label(self.f2, text="s")
        #s.grid(row=rownumber, column=3)
        rownumber += 1

        self.cv_end_volt.set("-100.0")
        s = Label(self.f2, text="End Volt [V]")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_end_volt)
        s.grid(row=rownumber, column=2)
        #s = Label(self.f2, text="V")
        #s.grid(row=rownumber, column=3)
        rownumber += 1

        self.cv_compliance.set("1.0")
        s = Label(self.f2, text="Compliance")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_compliance)
        s.grid(row=rownumber, column=2)
        self.cv_compliance_scale.set("uA")
        s = OptionMenu(self.f2, self.cv_compliance_scale, *compliance_choices)
        s.grid(row=rownumber, column=3)
        rownumber += 1

        self.cv_recipients.set("adapbot@gmail.com")
        s = Label(self.f2, text="Email data to:")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_recipients)
        s.grid(row=rownumber, column=2)
        rownumber += 1

        s = Label(self.f2, text="Agilent LCRMeter Parameters", relief=RAISED)
        s.grid(row=rownumber, column=1, columnspan=2)
        rownumber += 1

        self.cv_impedance = StringVar()
        s = Label(self.f2, text="Function")
        s.grid(row=rownumber, column=1)
        function_choices = {
            "CPD",
            "CPQ",
            "CPG",
            "CPRP",
            "CSD",
            "CSQ",
            "CSRS",
            "LPD",
            "LPQ",
            "LPG",
            "LPRP",
            "LPRD",
            "LSD",
            "LSQ",
            "LSRS",
            "LSRD",
            "RX",
            "ZTD",
            "ZTR",
            "GB",
            "YTD",
            "YTR",
            "VDID",
        }
        self.cv_function_choice = StringVar()
        self.cv_function_choice.set("CPRP")
        s = OptionMenu(self.f2, self.cv_function_choice, *function_choices)
        s.grid(row=rownumber, column=2)
        rownumber += 1

        self.cv_impedance.set("10000")
        s = Label(self.f2, text="Impedance")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_impedance)
        s.grid(row=rownumber, column=2)
        s = Label(self.f2, text="Ω")
        s.grid(row=rownumber, column=3)
        rownumber += 1

        self.cv_frequencies.set("10000")
        s = Label(self.f2, text="Frequencies")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_frequencies)
        s.grid(row=rownumber, column=2)
        s = Label(self.f2, text="Hz")
        s.grid(row=rownumber, column=3)
        rownumber += 1

        self.cv_amplitude.set("0.2")
        s = Label(self.f2, text="Signal Amplitude")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_amplitude)
        s.grid(row=rownumber, column=2)
        s = Label(self.f2, text="V")
        s.grid(row=rownumber, column=3)
        rownumber += 1

        cv_int_choices = {"Short", "Medium", "Long"}
        s = Label(self.f2, text="Integration time")
        s.grid(row=rownumber, column=1)
        self.cv_integration.set("Short")
        s = OptionMenu(self.f2, self.cv_integration, *cv_int_choices)
        s.grid(row=rownumber, column=2)
        rownumber += 1

        self.cv_source_choice.set("Keithley 2657a")
        s = OptionMenu(self.f2, self.cv_source_choice, *source_choices)
        s.grid(row=0, column=5)

        self.cv_canvas = FigureCanvasTkAgg(self.cv_f, master=self.f2)
        self.cv_canvas.get_tk_widget().grid(row=rownumber, column=0, columnspan=10)
        self.cv_a.set_title("CV")
        self.cv_a.set_xlabel("Voltage")
        self.cv_a.set_ylabel("Capacitance")
        self.cv_canvas.draw()
        rownumber += 1

        s = Label(self.f2, text="Progress:")
        s.grid(row=rownumber, column=1)

        self.cv_pb = ttk.Progressbar(
            self.f2, orient="horizontal", length=200, mode="determinate"
        )
        self.cv_pb.grid(row=rownumber, column=2, columnspan=5)
        self.cv_pb["maximum"] = 100
        self.cv_pb["value"] = 0
        rownumber += 1

        s = Label(self.f2, text="Est finish at:")
        s.grid(row=rownumber, column=1)
        cv_timetext = str(time.asctime(time.localtime(time.time()))) #fix ending time
        self.timer = Label(self.f2, text=cv_timetext)
        self.timer.grid(row=rownumber, column=2)

        rownumber += 1

        self.cv_instability.set("1.0")
        s = Label(self.f2, text="Instablity (%)")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_instability)
        s.grid(row=rownumber, column=2)
        rownumber += 1

        self.cv_instability_wait_time.set("1.0")
        s = Label(self.f2, text="Instability Wait Time (s)")
        s.grid(row=rownumber, column=1)
        s = Entry(self.f2, textvariable=self.cv_instability_wait_time)
        s.grid(row=rownumber, column=2)

        s = Button(self.f2, text="Start CV", command=self.cv_prepare_values)
        s.grid(row=7, column=7)

        s = Button(self.f2, text="Stop", command=self.quit)
        s.grid(row=8, column=7)

        print("finished drawing")

        """
        Multiple IV GUI
        """

        if "Windows" in platform.platform():
            self.multiv_filename.set("iv_data")
            s = Label(self.f4, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(self.f4, textvariable=self.multiv_filename)
            s.grid(row=0, column=2)

        s = Label(self.f4, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f4, textvariable=self.multiv_start_volt)
        s.grid(row=1, column=2)
        s = Label(self.f4, text="V")
        s.grid(row=1, column=3)

        s = Label(self.f4, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f4, textvariable=self.multiv_end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f4, text="V")
        s.grid(row=2, column=3)

        s = Label(self.f4, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f4, textvariable=self.multiv_step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f4, text="V")
        s.grid(row=3, column=3)

        s = Label(self.f4, text="Repeat Times")
        s.grid(row=4, column=1)
        s = Entry(self.f4, textvariable=self.multiv_times)
        s.grid(row=4, column=2)

        s = Label(self.f4, text="Hold Time")
        s.grid(row=5, column=1)
        s = Entry(self.f4, textvariable=self.multiv_hold_time)
        s.grid(row=5, column=2)
        s = Label(self.f4, text="s")
        s.grid(row=5, column=3)

        s = Label(self.f4, text="Compliance")
        s.grid(row=6, column=1)
        s = Entry(self.f4, textvariable=self.multiv_compliance)
        s.grid(row=6, column=2)
        self.multiv_compliance_scale.set("uA")
        s = OptionMenu(self.f4, self.multiv_compliance_scale, *compliance_choices)
        s.grid(row=6, column=3)

        self.multiv_recipients.set("adapbot@gmail.com")
        s = Label(self.f4, text="Email data to:")
        s.grid(row=7, column=1)
        s = Entry(self.f4, textvariable=self.multiv_recipients)
        s.grid(row=7, column=2)

        source_choices = {"Keithley 2400", "Keithley 2657a"}
        self.multiv_source_choice.set("Keithley 2657a")
        s = OptionMenu(self.f4, self.multiv_source_choice, *source_choices)
        s.grid(row=0, column=7)

        s = Label(self.f4, text="Progress:")
        s.grid(row=11, column=1)

        s = Label(self.f4, text="Est finish at:")
        s.grid(row=12, column=1)

        self.multiv_timer = Label(self.f4, text=timetext)
        self.multiv_timer.grid(row=12, column=2)

        self.multiv_pb = ttk.Progressbar(
            self.f4, orient="horizontal", length=200, mode="determinate"
        )
        self.multiv_pb.grid(row=11, column=2, columnspan=5)
        self.multiv_pb["maximum"] = 100
        self.multiv_pb["value"] = 0

        self.multiv_canvas = FigureCanvasTkAgg(self.multiv_f, master=self.f4)
        self.multiv_canvas.get_tk_widget().grid(row=8, columnspan=10)
        self.multiv_a.set_title("IV")
        self.multiv_a.set_xlabel("Voltage")
        self.multiv_a.set_ylabel("Current")

        self.multiv_canvas.draw()

        s = Button(self.f4, text="Start IVs", command=self.multiv_prepare_values)
        s.grid(row=3, column=7)

        s = Button(self.f4, text="Stop", command=self.quit)
        s.grid(row=4, column=7)

        """
        Current Monitor IV
        """

        if "Windows" in platform.platform():
            self.curmon_filename.set("iv_data")
            s = Label(self.f5, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(self.f5, textvariable=self.curmon_filename)
            s.grid(row=0, column=2)

        s = Label(self.f5, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f5, textvariable=self.curmon_start_volt)
        s.grid(row=1, column=2)
        s = Label(self.f5, text="V")
        s.grid(row=1, column=3)

        s = Label(self.f5, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f5, textvariable=self.curmon_end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f5, text="V")
        s.grid(row=2, column=3)

        s = Label(self.f5, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f5, textvariable=self.curmon_step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f5, text="V")
        s.grid(row=3, column=3)

        s = Label(self.f5, text="Test Time")
        s.grid(row=4, column=1)
        s = Entry(self.f5, textvariable=self.curmon_time)
        s.grid(row=4, column=2)
        s = Label(self.f5, text="M")
        s.grid(row=4, column=3)

        s = Label(self.f5, text="Hold Time")
        s.grid(row=5, column=1)
        s = Entry(self.f5, textvariable=self.curmon_hold_time)
        s.grid(row=5, column=2)
        s = Label(self.f5, text="s")
        s.grid(row=5, column=3)

        s = Label(self.f5, text="Compliance")
        s.grid(row=6, column=1)
        s = Entry(self.f5, textvariable=self.curmon_compliance)
        s.grid(row=6, column=2)
        self.curmon_compliance_scale.set("uA")
        s = OptionMenu(self.f5, self.curmon_compliance_scale, *compliance_choices)
        s.grid(row=6, column=3)

        self.curmon_recipients.set("adapbot@gmail.com")
        s = Label(self.f5, text="Email data to:")
        s.grid(row=7, column=1)
        s = Entry(self.f5, textvariable=self.curmon_recipients)
        s.grid(row=7, column=2)

        source_choices = {"Keithley 2400", "Keithley 2657a"}
        self.curmon_source_choice.set("Keithley 2657a")
        s = OptionMenu(self.f5, self.curmon_source_choice, *source_choices)
        s.grid(row=0, column=7)

        s = Label(self.f5, text="Progress:")
        s.grid(row=11, column=1)

        s = Label(self.f5, text="Est finish at:")
        s.grid(row=12, column=1)

        self.curmon_timer = Label(self.f5, text=timetext)
        self.curmon_timer.grid(row=12, column=2)

        self.curmon_pb = ttk.Progressbar(
            self.f5, orient="horizontal", length=200, mode="determinate"
        )
        self.curmon_pb.grid(row=11, column=2, columnspan=5)
        self.curmon_pb["maximum"] = 100
        self.curmon_pb["value"] = 0

        self.curmon_canvas = FigureCanvasTkAgg(self.curmon_f, master=self.f5)
        self.curmon_canvas.get_tk_widget().grid(row=8, columnspan=10)
        self.curmon_a.set_title("IV")
        self.curmon_a.set_xlabel("Voltage")
        self.curmon_a.set_ylabel("Current")

        self.curmon_canvas.draw()

        s = Button(self.f5, text="Start CurMon", command=self.curmon_prepare_values)
        s.grid(row=3, column=7)

        s = Button(self.f5, text="Stop", command=self.quit)
        s.grid(row=4, column=7)

        loadSettings(self)

    def update(self):
        while self.outputdata.qsize():
            try:
                #print("in update(self): ")
                (data, percent, timeremain) = self.outputdata.get(0)
                #print(data)
                if self.type is 0:
                    print("Percent done:" + str(percent))
                    self.pb["value"] = percent
                    self.pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        (line,) = self.a.plot(
                            map(lambda x: x * -1.0, voltages),
                            map(lambda x: x * -1.0, currents),
                        )
                    else:
                        (line,) = self.a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color("r")
                    self.a.set_title("IV")
                    self.a.set_xlabel("Voltage [V]")
                    self.a.set_ylabel("Current [A]")
                    self.canvas.draw()

                    timetext = str(
                        time.asctime(time.localtime(time.time() + timeremain))
                    )
                    self.timer = Label(self.f1, text=timetext)
                    self.timer.grid(row=12, column=2)

                elif self.type is 1:
                    (voltages, caps) = data
                    print("Percent done:" + str(percent))
                    self.cv_pb["value"] = percent
                    self.cv_pb.update()
                    # print "Caps:+++++++"
                    # print caps
                    # print "============="
                    colors = {0: "b", 1: "g", 2: "r", 3: "c", 4: "m", 5: "k"}
                    i = 0
                    for c in caps:
                        """
                        print "VOLTS++++++"
                        print voltages
                        print "ENDVOLTS===="
                        #(a, b) = c[0]
                        print "CAPSENSE+++++"
                        print c

                        print "ENDCAP======="
                        """

                        if self.first:

                            (line,) = self.cv_a.plot(
                                voltages,
                                c,
                                label=(self.cv_frequencies.get().split(",")[i] + "Hz"),
                            )
                            self.cv_a.legend()
                        else:#single frequency
                            (line,) = self.cv_a.plot(voltages, c)
                        line.set_antialiased(True)
                        line.set_color(colors.get(i))
                        i += 1
                        self.cv_a.set_title("CV")
                        self.cv_a.set_xlabel("Voltage [V]")
                        self.cv_a.set_ylabel("Capacitance [F]")
                        self.cv_canvas.draw()

                    totalpoints = (float(self.cv_start_volt2.get())-float(self.cv_start_volt1.get()))/float(self.cv_step_volt1.get()) + (float(self.cv_start_volt3.get())-float(self.cv_start_volt2.get()))/float(self.cv_step_volt2.get()) + (float(self.cv_end_volt.get())-float(self.cv_start_volt3.get()))/float(self.cv_step_volt3.get())

                    cv_total_time = totalpoints*(float(self.cv_instability_wait_time.get())+float({"Short": 0, "Medium": 1, "Long": 2}.get(self.cv_integration.get())))+(float(self.cv_start_volt2.get())-float(self.cv_start_volt1.get()))*float(self.cv_hold_time1.get())/float(self.cv_step_volt1.get()) + (float(self.cv_start_volt3.get())-float(self.cv_start_volt2.get()))*float(self.cv_hold_time2.get())/float(self.cv_step_volt2.get()) + (float(self.cv_end_volt.get())-float(self.cv_start_volt3.get()))*float(self.cv_hold_time3.get())/float(self.cv_step_volt3.get())
                    cv_total_time = -1*cv_total_time
                    global cv_start_time
                    timetext = str(
                        time.asctime(time.localtime(cv_start_time + cv_total_time))
                    )
                    self.timer = Label(self.f2, text=timetext)
                    self.timer.grid(row=23, column=2)
                    self.first = False

                elif self.type is 2:
                    pass

                elif self.type is 3:
                    if self.first:
                        # self.multiv_f.clf()
                        pass
                    print("Percent done:" + str(percent))
                    self.multiv_pb["value"] = percent
                    self.multiv_pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        (line,) = self.multiv_a.plot(
                            map(lambda x: x * -1.0, voltages),
                            map(lambda x: x * -1.0, currents),
                        )
                    else:
                        (line,) = self.multiv_a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color("r")
                    self.multiv_a.set_title("IV")
                    self.multiv_a.set_xlabel("Voltage [V]")
                    self.multiv_a.set_ylabel("Current [A]")
                    self.multiv_canvas.draw()

                    timetext = str(
                        time.asctime(time.localtime(time.time() + timeremain))
                    )
                    self.multiv_timer = Label(self.f4, text=timetext)
                    self.multiv_timer.grid(row=12, column=2)

                elif self.type is 4:

                    print("Percent done:" + str(percent))
                    self.curmon_pb["value"] = percent
                    self.curmon_pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        (line,) = self.curmon_a.plot(
                            map(lambda x: x * -1.0, voltages),
                            map(lambda x: x * -1.0, currents),
                        )
                    else:
                        (line,) = self.curmon_a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color("r")
                    self.curmon_a.set_title("IV")
                    self.curmon_a.set_xlabel("Voltage [V]")
                    self.curmon_a.set_ylabel("Current [A]")
                    self.curmon_canvas.draw()

                    timetext = str(
                        time.asctime(time.localtime(time.time() + timeremain))
                    )
                    self.curmon_timer = Label(self.f5, text=timetext)
                    self.curmon_timer.grid(row=12, column=2)
            except Queue.Empty:
                print("queue empty!")
                pass

    def quit(self):
        print("placing order")
        self.stop.put("random")
        self.stop.put("another random value")

    def prepare_values(self):
        print("preparing iv values")
        input_params = (
            (
                self.compliance.get(),
                self.compliance_scale.get(),
                self.start_volt.get(),
                self.end_volt.get(),
                self.step_volt.get(),
                self.hold_time.get(),
                self.source_choice.get(),
                self.recipients.get(),
                self.filename.get(),
            ),
            0,
        )
        self.inputdata.put(input_params)
        self.f.clf()
        self.a = self.f.add_subplot(111)
        self.type = 0

    def cv_prepare_values(self):
        print("preparing cv values")
        self.first = True
        global cv_start_time
        cv_start_time = time.time()

        # Protection for new CV box
        t_cv_end_volt = 0.0
        # print self.cv_end_volt.get()
        if abs(float(self.cv_end_volt.get())) < 200:
            t_cv_end_volt = float(self.cv_end_volt.get())
        else:
            print("NEVER go over 200V! setting limit to 200V")
            t_cv_end_volt = (
                200 * float(self.cv_end_volt.get()) / abs(float(self.cv_end_volt.get()))
            )
        # print t_cv_end_volt
        t_s_cv_end_volt = str(t_cv_end_volt)
        # print t_s_cv_end_volt
        input_params = ((
            self.cv_compliance.get(),
            self.cv_compliance_scale.get(),
            
            self.cv_start_volt1.get(),
            self.cv_step_volt1.get(),
            self.cv_hold_time1.get(),
            self.cv_start_volt2.get(),
            self.cv_step_volt2.get(),
            self.cv_hold_time2.get(),
            self.cv_start_volt3.get(),
            self.cv_step_volt3.get(),
            self.cv_hold_time3.get(),

            t_s_cv_end_volt,
            self.cv_source_choice.get(),
            list(map(lambda x: x.strip(), self.cv_frequencies.get().split(","))),
            self.cv_function_choice.get(),
            self.cv_amplitude.get(),
            self.cv_impedance.get(),
            self.cv_integration.get(),
            self.cv_recipients.get(),
            self.cv_filename.get(),
            self.cv_instability.get(),
            self.cv_instability_wait_time.get(),
            self.cv_temperature.get(),
            self.cv_humidity.get(),
            self.cv_your_name.get(),
            self.cv_vendor.get(),
            self.cv_product_name.get(),
            self.cv_Type.get(),
            self.cv_geometry.get(),
            self.cv_wafer.get(),
            self.cv_sensor_type.get(),
            self.cv_chip_num.get(),
            self.cv_fluence.get(),
            self.cv_fluence_type.get(),
            self.cv_area.get(),
            self.cv_set_number.get()
            ),1)
        
        saveSettings(self)
        self.inputdata.put(input_params)
        self.cv_f.clf()
        self.cv_a = self.cv_f.add_subplot(111)
        self.type = 1

    def multiv_prepare_values(self):

        print("preparing mult iv values")
        self.first = True
        input_params = (
            (
                self.multiv_compliance.get(),
                self.multiv_compliance_scale.get(),
                self.multiv_start_volt.get(),
                self.multiv_end_volt.get(),
                self.multiv_step_volt.get(),
                self.multiv_hold_time.get(),
                self.multiv_source_choice.get(),
                self.multiv_recipients.get(),
                self.multiv_filename.get(),
                self.multiv_times.get(),
            ),
            3,
        )
        self.inputdata.put(input_params)
        self.multiv_f.clf()
        self.multiv_a = self.multiv_f.add_subplot(111)
        self.type = 3

    def curmon_prepare_values(self):

        print("preparing current monitor values")
        self.first = True
        input_params = (
            (
                self.curmon_compliance.get(),
                self.curmon_compliance_scale.get(),
                self.curmon_start_volt.get(),
                self.curmon_end_volt.get(),
                self.curmon_step_volt.get(),
                self.curmon_hold_time.get(),
                self.curmon_source_choice.get(),
                self.curmon_recipients.get(),
                self.curmon_filename.get(),
                self.curmon_time.get(),
            ),
            4,
        )
        self.inputdata.put(input_params)
        self.curmon_f.clf()
        self.curmon_a = self.curmon_f.add_subplot(111)
        self.type = 4


def getvalues(input_params, dataout, stopqueue):
    if "Windows" in platform.platform():
        (
            compliance,
            compliance_scale,
            start_volt,
            end_volt,
            step_volt,
            hold_time,
            source_choice,
            recipients,
            filename,
        ) = input_params
    else:
        (
            compliance,
            compliance_scale,
            start_volt,
            end_volt,
            step_volt,
            hold_time,
            source_choice,
            recipients,
            thowaway,
        ) = input_params
        filename = tkFileDialog.asksaveasfilename(
            initialdir="~",
            title="Save data",
            filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")),
        )
    print("File done")

    try:
        comp = float(
            float(compliance)
            * ({"mA": 1e-3, "uA": 1e-6, "nA": 1e-9}.get(compliance_scale, 1e-6))
        )
        source_params = (
            int(float(start_volt)),
            int(float(end_volt)),
            (float(step_volt)),
            float(hold_time),
            comp,
        )
    except ValueError:
        print("Please fill in all fields!")
    data = ()
    if source_params is None:
        pass
    else:
        print(source_choice)
        choice = 0
        if "2657a" in source_choice:
            print("asdf keithley 366")
            choice = 1
        data = GetIV(source_params, choice, dataout, stopqueue)

    fname = (
        (
            filename + "_" + str(time.asctime(time.localtime(time.time()))) + ".xlsx"
        ).replace(" ", "_")
    ).replace(":", "_")

    data_out = xlsxwriter.Workbook(fname)
    if "Windows" in platform.platform():
        fname = "./" + fname
    worksheet = data_out.add_worksheet()

    (v, i) = data
    values = []
    pos = v[0] > v[1]
    for x in np.arange(0, len(v), 1):
        values.append((v[x], i[x]))
    row = 0
    col = 0

    chart = data_out.add_chart({"type": "scatter", "subtype": "straight_with_markers"})

    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col + 1, cur)
        row += 1

    chart.add_series(
        {
            "categories": "=Sheet1!$A$1:$A$" + str(row),
            "values": "=Sheet1!$B$1:$B$" + str(row),
        }
    )
    chart.set_x_axis(
        {
            "name": "Voltage [V]",
            "major_gridlines": {"visible": True},
            "minor_tick_mark": "cross",
            "major_tick_mark": "cross",
            "line": {"color": "black"},
            "reverse": pos,
        }
    )
    chart.set_y_axis(
        {
            "name": "Current [A]",
            "major_gridlines": {"visible": True},
            "minor_tick_mark": "cross",
            "major_tick_mark": "cross",
            "line": {"color": "black"},
            "reverse": pos,
        }
    )
    chart.set_legend({"none": True})
    worksheet.insert_chart("D2", chart)
    data_out.close()

    try:
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())

        print(sentTo)
        print(fname)
        sendMail(fname, sentTo)
    except:
        pass


def cv_getvalues(input_params, dataout, stopqueue):
    #print("in cv_getvalues: " + str(input_params))
    if "Windows" in platform.platform():
        (
            compliance,
            compliance_scale,
            start_volt1,
            step_volt1,
            hold_time1,
            start_volt2,
            step_volt2,
            hold_time2,
            start_volt3,
            step_volt3,
            hold_time3,
            end_volt,
            source_choice,
            frequencies,
            function,
            amplitude,
            impedance,
            integration,
            recipients,
            filename,
            instability,
            instability_wait_time,
            temperature,
            humidity,
            your_name,
            vendor,
            product_name,
            Type,
            geometry,
            wafer,
            sensor_type,
            chip_num,
            fluence,
            fluence_type,
            area,
            set_number
        ) = input_params
        filename = "./" + filename
        
    else:
        raise Exception("WRONG SETUP call rick!")
        (
            compliance,
            compliance_scale,
            start_volt1,
            step_volt1,
            hold_time1,
            start_volt2,
            step_volt2,
            hold_time2,
            start_volt3,
            step_volt3,
            hold_time3,
            end_volt,
            source_choice,
            frequencies,
            function,
            amplitude,
            impedance,
            integration,
            recipients,
            thowaway,
        ) = input_params
        filename = tkFileDialog.asksaveasfilename(
            initialdir="~",
            title="Save data",
            filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")),
        )

    try:
        # step_volt was originally int(float(step_volt)), but it was causing problems with steps sizes < 1.0.  Since it is 'scaled'
        # later on, it doesn't need to be cast as an int first.
        comp = float(
            float(compliance)
            * ({"mA": 1e-3, "uA": 1e-6, "nA": 1e-9}.get(compliance_scale, 1e-6))
        )
        params = (
            int(float(start_volt1)),
            float(step_volt1),
            float(hold_time1),
            int(float(start_volt2)),
            float(step_volt2),
            float(hold_time2),
            int(float(start_volt3)),
            float(step_volt3),
            float(hold_time3),
            int(float(end_volt)),
            
            comp,
            frequencies,
            float(amplitude),
            function,
            int(impedance),
            {"Short": 0, "Medium": 1, "Long": 2}.get(integration),
            float(instability) / 100,
            float(instability_wait_time),
            str(temperature),
            str(humidity),
            str(your_name),
            vendor,
            str(product_name),
            str(Type),
            geometry,
            str(wafer),
            str(sensor_type),
            str(chip_num),
            str(fluence),
            fluence_type,
            str(area),
            str(set_number)
        )

        #print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + params + "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        #print(params)
    except ValueError as val:
        print("Your exception is {}".format(val))
        #params = None
        print("Please fill in all fields for CV!")
        pass
    except Exception as e:
        print(("catch {}".format(e)))
        print(e)
        pass
    data = ()
    if params is None:
        pass
    else:
        data = GetCV(params, {"Keithley 2657a":1, "Keithley 2400":0}.get(source_choice), dataout, stopqueue)
        fname = (((filename+"_"+str(time.asctime(time.localtime(time.time())))+".xlsx").replace(" ", "_")).replace(":", "_"))


    data_out = xlsxwriter.Workbook(fname)
    if "Windows" in platform.platform():
        fname = "./" + fname
        
    if geometry == "single_pad":
        if fluence_type == "Non-irradiated":
            seriesName = vendor+ "_"+ Type+"_"+ wafer+"_"+ set_number+"_" +sensor_type+"_"+ chip_num
        else:
            seriesName = vendor+"_"+ Type+"_"+ wafer+"_"+ set_number+"_"+ sensor_type+"_"+ chip_num + "_"+ fluence+"_"+ fluence_type
    else:
        if fluence_type == "Non-irradiated":
            seriesName = vendor+"_"+ Type+"_"+ geometry+"_"+ wafer+"_"+ set_number+"_"+ sensor_type+"_"+ chip_num
        else:
            seriesName = vendor+"_"+ Type+"_"+ geometry+"_"+ wafer+"_"+ set_number+"_"+ sensor_type+"_"+ chip_num + "_"+ fluence+"_"+ fluence_type
   
    worksheet = data_out.add_worksheet(seriesName)
    green_cell = data_out.add_format()
    green_cell.set_font_color('green')
    green_cell.set_bg_color('#acff9b')
    green_cell.set_bold()
    cell_format = data_out.add_format()
    cell_format.set_bold()
    
    #log.critical(data)
    #log.critical(type(data))
    
    (v,i,c,r) =data
    
    '''The single voltage range version outputs are in form of
    v = [], i =[[]], c=[[]],r=[[]] for single frequency, and
    v = [], i =[[],[],...], c =[[],[],...], r =[[],[],...] for multiple fequency, length of i,c, or r is n for n frequencies

    However, in this m multiple voltage ranges, v is still v =[] but the length of i,c,or r is m*n.
    So, we need to re-arrange the i,c,r lists
    '''

    v = v[0] + v[1] +v[2]
    i = i[0] + i[1] +i[2]
    c = c[0] + c[1] +c[2]
    r = r[0] + r[1] +r[2]
    i = [i[0] + i[1] +i[2]]
    c = [c[0] + c[1] +c[2]]
    r = [r[0] + r[1] +r[2]]
        

    #print(v)
    #print(i)
    #print(c)
    #print(r)

#===============================================================================

    #here are the codes that write the data into an excel file
    row = 9
    col = 0

    totalRows = row + len(v)
    
    chart = data_out.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
    worksheet.write(8, 0, "V",cell_format)
    

    for volt in v:
        worksheet.write(row, col, volt)
        row += 1

    col += 1
    last_col = col

    for f in frequencies:
        worksheet.write(7, col, "Freq=" + f + "Hz")
        col += 3

    col = last_col
    row = 9
    for frequency in i:
        worksheet.write(8, col, "I",cell_format)
        row = 9
        for current in frequency:
            worksheet.write(row, col, current)
            row += 1
        col += 3

    col = last_col + 1
    last_col = col
    for frequency in c:
        worksheet.write(8, col, "C",cell_format)
        row = 9
        for cap in frequency:
            worksheet.write(row, col, cap)
            row += 1
        col += 3

    col = last_col + 1
    last_col = col

    fs = 0
    for frequency in r:
        fs += 1
        worksheet.write(8, col, "R",cell_format)
        row = 9
        for res in frequency:
            worksheet.write(row, col, res)
            row += 1
        col += 3

    row += 5

    worksheet.write('A1',seriesName)
    worksheet.write('H4',"Area")
    worksheet.write('I4'," = ")
    worksheet.write('J4',area)
    worksheet.write('F1',"Name")
    worksheet.write('G1',your_name)
    worksheet.write('A2',"Temperature")
    worksheet.write('B2',temperature)
    worksheet.write('C2',"Humidity")
    worksheet.write('D2',humidity)

    worksheet.write('A3',vendor)
    worksheet.write('B3',product_name)
    worksheet.write('C3',Type)
    worksheet.write('D3',geometry)
    worksheet.write('E3',"Wafer " +str(wafer))
    worksheet.write('F3',sensor_type)
    worksheet.write('G3',str(chip_num))
    worksheet.write('H3',fluence)
    worksheet.write('I3',fluence_type)
    
    new_row = 8
    oneOverCsq_ = []
    negV_ = []
    negI_ = []
    CdivA = []
    depth = []
    AsqvCsq =[]
    derivative = []
    doping = []
    
    analyzed_param_names = ['V','I','C','R','1/C^2','-V','-I','C/A','Depth (um)','A^2/C^2 (cm^4/F^2)','Derivative (cm4/V·F²)','N (cm-3)']
    for colnum in range(len(analyzed_param_names)):
        if colnum == 8 or colnum == 9 or colnum == 10 or colnum == 11:
            worksheet.write(new_row,colnum, analyzed_param_names[colnum],green_cell)
        else:
            worksheet.write(new_row,colnum, analyzed_param_names[colnum],cell_format)
            
    for colnum in range(len(analyzed_param_names)):
        for rownum in range(new_row+1,totalRows):
            if colnum == 4:
                worksheet.write(rownum,colnum, '=1/(C{}^2)'.format(rownum+1)) #1/C^2
            elif colnum == 5:
                worksheet.write(rownum,colnum,'=ABS(A{})'.format(rownum+1)) #-V
            elif colnum == 6:
                worksheet.write(rownum,colnum,'=ABS(B{})'.format(rownum+1)) #-I
            elif colnum == 7:
                worksheet.write(rownum,colnum,'=C{}'.format(rownum+1)+'/$J$4') #C/A
            elif colnum == 8:
                worksheet.write(rownum,colnum,'=11.9*0.0000000000000885*10000/C{}*$J$4'.format(rownum+1)) #depth
            elif colnum == 9:
                worksheet.write(rownum,colnum,'=($J$4/C{})^2'.format(rownum+1)) #A^2/C^2
            elif colnum == 10:
                worksheet.write(rownum,colnum,'=(J{}'.format(rownum+2) +'-J{})/'.format(rownum+1) +'(F{}-'.format(rownum+2) + 'F{})'.format(rownum+1))
            elif colnum == 11:
                worksheet.write(rownum,colnum, '=2/(1.6E-19*11.9*0.0000000000000885*K{})'.format(rownum+1))

            

#===============================================================================
    #-------graphing Excel charts-------------
    endingNum = len(v)-1 #doping and derivative has 1 less than voltage
    rw=8
    #C vs V
    makeExcelChart(data_out, worksheet, seriesName, 'O3', 'scatter', rw +2, endingNum+rw+2,'Bias Voltage [V]','Capacitance [F]',
                    False, seriesName+ " C vs V",'F','C',abs(max(v)),abs(min(v)),0,3e-10)
    makeExcelChart(data_out, worksheet, seriesName, 'W3', 'scatter', rw +2, endingNum+rw+2,'Bias Voltage [V]','Capacitance [F]',
                    True, seriesName+ " C vs V",'F','C',abs(max(v)),abs(min(v)),1e-12,3e-10)
    #1/C^2 vs V
    makeExcelChart(data_out, worksheet, seriesName, 'O19', 'scatter', rw +2, endingNum+rw+2,'Bias Voltage [V]','1/C^2 [1/F^2]',
                    True,seriesName+ " 1/C^2 vs V",'F','E',abs(max(v)),abs(min(v)),1e19,1e23)
    makeExcelChart(data_out, worksheet, seriesName, 'W19', 'scatter', rw +2, endingNum+rw+2,'Bias Voltage [V]','1/C^2 [1/F^2]',
                    True,seriesName+ " 1/C^2 vs V",'F','E',40,70,1e20,1e23)

    #doping vs V
    makeExcelChart(data_out, worksheet, seriesName, 'O35', 'scatter', rw +2,endingNum+rw+2,'Bias Voltage [V]','Doping profile [N/cm^3]',
                    True,seriesName+ " Doping vs V",'F','L',abs(max(v)),abs(min(v)),1e12,1e17)
    makeExcelChart(data_out, worksheet, seriesName, 'W35', 'scatter', rw +2,endingNum+rw+2,'Bias Voltage [V]','Doping profile [N/cm^3]',
                    True,seriesName+ " Doping vs V",'F','L',40,70,1e12,1e17)
    #doping vs depth
    makeExcelChart(data_out, worksheet, seriesName, 'O51', 'scatter', rw +2, endingNum+rw+2,'Depth [um]','Doping profile [N/cm^3]',
                    True,seriesName+ " Doping vs Depth",'I','L',0,50,1e12,1e17)
    makeExcelChart(data_out, worksheet, seriesName, 'W51', 'scatter', rw +2, endingNum+rw+2,'Depth [um]','Doping profile [N/cm^3]',
                    True,seriesName+ " Doping vs Depth",'I','L',1,5,1e12,1e17)

    data_out.close()

    try:
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())

        print(sentTo)
        sendMail(fname, sentTo)
    except Exception as e:
        print(e)
        print("Failed to get recipients")
        pass


def makeExcelChart(workbook, sheet, seriesName, chartPosition, chartType, startingCol, endingCol, 
                  xLabel, yLabel, yLog,  chartTitle, xcol, ycol, xlimmin,xlimmax,ylimmin,ylimmax):
    chart = workbook.add_chart({'type': chartType})
    chart.add_series({
    'name': seriesName,
    'categories': '=\'{}\'!$'.format(seriesName) + '{}$'. format(xcol) +'{}:$'.format(startingCol)+ '{}$'.format(xcol)+'{}'.format(endingCol),
    'values':  '=\'{}\'!$'.format(seriesName) + '{}$'. format(ycol) +'{}:$'.format(startingCol)+ '{}$'.format(ycol)+'{}'.format(endingCol),
    'marker': {'type': 'circle','size': 5, 'fill': {'none':True}},
    'line': {'dash_type': 'dash', 'width': 0.3}
    })
    chart.set_x_axis({'name': xLabel, 
                      'major_tick_mark': 'inside',
                      'minor_tick_mark': 'inside',
                      'min': xlimmin,
                      'max': xlimmax
                     })
    if yLog == True:
        chart.set_y_axis({'log_base': 10, 'name': yLabel, 
                      'major_gridlines': {'visible': True,'line': {'width': 0.15}},
                      'min': ylimmin,
                      'max': ylimmax,
                      'major_tick_mark': 'inside',
                      'minor_tick_mark': 'inside'
                         })
    else:
        chart.set_y_axis({'name': yLabel, 
                      'major_gridlines': {'visible': True,'line': {'width': 0.15}},
                      'min': ylimmin,
                      'max': ylimmax,
                      'major_tick_mark': 'inside',
                      'minor_tick_mark': 'inside'
                         })
        
    chart.set_title({'name':chartTitle,
                    'name_font':  {'name': 'Arial', 'size': 9,'bold': True}})
    
    chart.set_legend({'position': 'bottom'})
    sheet.insert_chart(chartPosition,chart)


# TODO: Implement value parsing from gui
def spa_getvalues(input_params, dataout):
    pass


def multiv_getvalues(input_params, dataout, stopqueue):
    if "Windows" in platform.platform():
        (
            compliance,
            compliance_scale,
            start_volt,
            end_volt,
            step_volt,
            hold_time,
            source_choice,
            recipients,
            filename,
            times_str,
        ) = input_params
    else:
        (
            compliance,
            compliance_scale,
            start_volt,
            end_volt,
            step_volt,
            hold_time,
            source_choice,
            recipients,
            thowaway,
            times_str,
        ) = input_params
        filename = tkFileDialog.asksaveasfilename(
            initialdir="~",
            title="Save data",
            filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")),
        )
    print("File done")

    try:
        comp = float(
            float(compliance)
            * ({"mA": 1e-3, "uA": 1e-6, "nA": 1e-9}.get(compliance_scale, 1e-6))
        )
        source_params = (
            int(float(start_volt)),
            int(float(end_volt)),
            (float(step_volt)),
            float(hold_time),
            comp,
        )
        times = int(times_str)

    except ValueError:
        print("Please fill in all fields!")
    data = ()

    while times > 0:
        if not stopqueue.empty():
            break

        if source_params is None:
            pass
        else:
            print(source_choice)
            choice = 0
            if "2657a" in source_choice:
                print("asdf keithley 366")
                choice = 1
            data = GetIV(source_params, choice, dataout, stopqueue)
        fname = (
            (
                filename
                + "_"
                + str(time.asctime(time.localtime(time.time())))
                + ".xlsx"
            ).replace(" ", "_")
        ).replace(":", "_")
        print(fname)
        data_out = xlsxwriter.Workbook(fname)
        if "Windows" in platform.platform():
            fname = "./" + fname
        worksheet = data_out.add_worksheet()

        (v, i) = data
        values = []
        pos = v[0] > v[1]
        for x in np.arange(0, len(v), 1):
            values.append((v[x], i[x]))
        row = 0
        col = 0
        chart = data_out.add_chart(
            {"type": "scatter", "subtype": "straight_with_markers"}
        )
        for volt, cur in values:
            worksheet.write(row, col, volt)
            worksheet.write(row, col + 1, cur)
            row += 1
        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$" + str(row),
                "values": "=Sheet1!$B$1:$B$" + str(row),
                "marker": {"type": "triangle"},
            }
        )
        chart.set_x_axis(
            {
                "name": "Voltage [V]",
                "major_gridlines": {"visible": True},
                "minor_tick_mark": "cross",
                "major_tick_mark": "cross",
                "line": {"color": "black"},
                "reverse": pos,
            }
        )
        chart.set_y_axis(
            {
                "name": "Current [A]",
                "major_gridlines": {"visible": True},
                "minor_tick_mark": "cross",
                "major_tick_mark": "cross",
                "line": {"color": "black"},
                "reverse": pos,
            }
        )
        chart.set_legend({"none": True})
        worksheet.insert_chart("D2", chart)
        data_out.close()

        try:
            mails = recipients.split(",")
            sentTo = []
            for mailee in mails:
                sentTo.append(mailee.strip())

            print(sentTo)
            sendMail(fname, sentTo)
        except:
            pass
        data_out.close()
        times -= 1


# TODO: Implement value parsing from gui
def curmon_getvalues(input_params, dataout, stopqueue):

    if "Windows" in platform.platform():
        (
            compliance,
            compliance_scale,
            start_volt,
            end_volt,
            step_volt,
            hold_time,
            source_choice,
            recipients,
            filename,
            total_time,
        ) = input_params
        filename = (
            (
                filename
                + "_"
                + str(time.asctime(time.localtime(time.time())))
                + ".xlsx"
            ).replace(" ", "_")
        ).replace(":", "_")
    else:
        (
            compliance,
            compliance_scale,
            start_volt,
            end_volt,
            step_volt,
            hold_time,
            source_choice,
            recipients,
            thowaway,
            total_time,
        ) = input_params
        filename = tkFileDialog.asksaveasfilename(
            initialdir="~",
            title="Save data",
            filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")),
        )
    print("File done")

    try:
        comp = float(
            float(compliance)
            * ({"mA": 1e-3, "uA": 1e-6, "nA": 1e-9}.get(compliance_scale, 1e-6))
        )
        source_params = (
            int(float(end_volt)),
            float(step_volt),
            float(hold_time),
            comp,
            int(total_time),
        )
    except ValueError:
        print("Please fill in all fields!")
    data = ()
    if source_params is None:
        pass
    else:
        print(source_choice)
        choice = 0
        if "2657a" in source_choice:
            print("asdf keithley 366")
            choice = 1
        data = curmon(source_params, choice, dataout, stopqueue)

    data_out = xlsxwriter.Workbook(filename)
    if "Windows" in platform.platform():
        fname = "./" + filename
    path = filename
    worksheet = data_out.add_worksheet()

    (v, i) = data
    values = []
    pos = v[0] > v[1]
    for x in np.arange(0, len(v), 1):
        values.append((v[x], i[x]))
    row = 0
    col = 0

    chart = data_out.add_chart({"type": "scatter", "subtype": "straight_with_markers"})

    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col + 1, cur)
        row += 1

    chart.add_series(
        {
            "categories": "=Sheet1!$A$1:$A$" + str(row),
            "values": "=Sheet1!$B$1:$B$" + str(row),
        }
    )
    chart.set_x_axis(
        {
            "name": "Voltage [V]",
            "major_gridlines": {"visible": True},
            "minor_tick_mark": "cross",
            "major_tick_mark": "cross",
            "line": {"color": "black"},
            "reverse": pos,
        }
    )
    chart.set_y_axis(
        {
            "name": "Current [A]",
            "major_gridlines": {"visible": True},
            "minor_tick_mark": "cross",
            "major_tick_mark": "cross",
            "line": {"color": "black"},
            "reverse": pos,
        }
    )
    chart.set_legend({"none": True})
    worksheet.insert_chart("D2", chart)
    data_out.close()

    try:
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())

        print(sentTo)
        sendMail(path, sentTo[0])
    except:
        pass


class ThreadedProgram:
    def __init__(self, master):
        self.master = master
        #print(type(master))
        #print(self.master.__class__.__name__)
        self.inputdata = Queue.Queue()
        self.outputdata = Queue.Queue()
        self.stopqueue = Queue.Queue()
        print("Generating GUI")

        self.running = 1
        try:
            self.gui = GuiPart(master, self.inputdata, self.outputdata, self.stopqueue)
        except Exception as e:
            print(e)
            #raw_input()
            input()
        self.thread1 = threading.Thread(target=self.workerThread1)
        self.thread1.start()
        self.periodicCall()
        self.measuring = False

        self.master.protocol("WM_DELETE_WINDOW", self.endapp)

    def periodicCall(self):
        # print "Period"
        self.gui.update()
        if self.stopqueue.qsize() == 1:
            pass
            # self.stopqueue.get()
            # print "Exiting program"
            # import sys
            # self.master.destroy()
            # self.running = 0
            # sys.exit(0)

        self.master.after(200, self.periodicCall)

    def workerThread1(self):
        while self.running:
            if self.inputdata.empty() is False and self.measuring is False:
                self.measuring = True
                print("Instantiating Threads")
                params= self.inputdata.get()

                if params[1] is 0:
                    getvalues(params[0], self.outputdata, self.stopqueue)
                elif params[1] is 1:
                    cv_getvalues(params[0], self.outputdata, self.stopqueue)
                elif params[1] is 2:
                    spa_getvalues(params[0], self.outputdata, self.stopqueue)
                elif params[1] is 3:
                    multiv_getvalues(params[0], self.outputdata, self.stopqueue)
                elif params[1] is 4:
                    curmon_getvalues(params[0], self.outputdata, self.stopqueue)
                else:
                    pass
                self.measuring = False

    def endapp(self):
        self.running = 0
        self.master.destroy()
        import sys

        sys.exit(0)


if __name__ == "__main__":
    try:
        root = Tk()
        root.geometry("800x900")
        root.title("LabMaster GUI v5")
        client = ThreadedProgram(root)
        root.mainloop()

    except Exception as e:
        print("Hi," + e)
        input()
