import visa,time
from Instrument import Instrument
from numpy import linspace
class Keithley2400(object):

    def __init__(self, gpib=22):
        """Set up gpib controller for device"""

        assert(gpib >= 0), "Please enter a valid gpib address"
        self.gpib_addr = gpib

        print("Initializing keithley 2400")
        rm = visa.ResourceManager()
        self.inst = rm.open_resource(rm.list_resources()[0])
        for x in rm.list_resources():
            if str(self.gpib_addr) in x:
                print("found")
                self.inst = rm.open_resource(x)

        print(self.inst.query("*IDN?;"))
        self.inst.write("*RST;")
        # self.inst.query("*RST;*IDN?;")
        self.inst.write("*ESE 1;*SRE 32;*CLS;:FUNC:CONC ON;:FUNC:ALL;:TRAC:FEED:CONT NEV;:RES:MODE MAN;")
        print(self.inst.query(":SYST:ERR?;"))
        self.inst.write(":DISP:CND;")
        self.inst.timeout = 10000

    def configure_measurement(self, _meas=1, output_level=0, compliance=5e-6):
        """Set what we are measuring and autoranging"""

        meas = {0:":VOLT", 1:":CURR", 2:":RES"}.get(_meas, ":CURR")
        self.inst.write(meas + ":RANG:AUTO ON;")
        _src = 0
        self.__configure_source(_src, output_level, compliance)

    def __configure_source(self, _func=0, output_level=0, compliance=5e-6):
        """Set output level and function"""

        assert(_func >= 0 and _func < 2), "Invalid Source function"
        assert(output_level > -1100 and output_level < 1100), "Voltage out of range"
        assert(compliance < 0.5), "Compliance out of range"
        func = {0:"VOLT", 1:"CURR"}.get(_func, "VOLT")
        self.inst.write("SOUR:FUNC " + func + ";:SOUR:" + func + " " + str(float(output_level)) + ";:CURR:PROT " + str(float(compliance)) + ";")

    def set_output(self, level=0):
        self.inst.write(":SOUR:VOLT " + str(float(level)))
        self.enable_output(True)

    def enable_output(self, output=False):
        """Sets output of front panel"""

        if output is True:
            self.inst.write("OUTP ON;")
            return
        self.inst.write("OUTP OFF;")

    def __configure_multipoint(self, arm_count=1, trigger_count=1, mode=0):
        """Configures immediate triggering and arming"""

        assert(mode >= 0 and mode < 3), "Invalid mode"
        source_mode = {0:"FIX", 1:"SWE", 2:"LIST"}.get(mode)
        self.inst.write(":ARM:COUN " + str(arm_count) + ";:TRIG:COUN " + str(trigger_count) + ";:SOUR:VOLT:MODE " + source_mode + ";:SOUR:CURR:MODE " + source_mode)

    def __configure_trigger(self, delay=0.0):
        self.inst.write("ARM:SOUR IMM;:ARM:TIM 0.010000;:TRIG:SOUR IMM;:TRIG:DEL " + str(delay) + ";")

    def __initiate_trigger(self):
        self.inst.write(":TRIG:CLE;:INIT;")

    def __wait_operation_complete(self):

        self.inst.write("*OPC;")


    # TODO parse data from visa buffer
    def __fetch_measurements(self):
        read_bytes = self.inst.query(":FETC?")
        return float(read_bytes.split(",")[1])

    def get_current(self, delay=0):
        self.__configure_multipoint()
        self.__configure_trigger(delay)
        self.__initiate_trigger()
        self.__wait_operation_complete()
        return self.__fetch_measurements()

class Keithley2657a(Instrument):

    def __init__(self, gpib=24):
        self.inst=None
        self.connect("Keithley","2657a")

        #self.inst.write("smua.reset()")
        #self.inst.write("smua.measure.adc=smua.ADC_INTEGRATE")
        self.inst.write("smua.measure.nplc=25")
        #self.inst.write("smua.measure.count= 1000")
        #self.inst.write("smua.measure.interval= 10e-6")

        #self.inst.write("print(errorqueue.next())")
        self.inst.write("display.smua.measure.func = display.MEASURE_DCAMPS")
        self.inst.write("setup.recall(1)")
        self.inst.timeout = 10000

    def __configure_source(self, _source=1):
        source = {0:"OUTPUT_DCAMPS", 1:"OUTPUT_DCVOLTS"}.get(_source, "OUTPUT_DCVOLTS")
        self.inst.write("smua.source.func = smua." + source)

    def set_output(self, level=0):
        self.inst.write("smua.source.levelv = " + str(level))
        self.enable_output(True)

    def set_compliance(self, limit=0.1):
        self.inst.write("smua.source.limiti = " + str(limit))

    def __configure_compliance(self, limit=0.1):
        self.inst.write("smua.source.limiti = " + str(limit))
        self.inst.write("smua.measure.rangei = " + str(10*limit))

    def configure_measurement(self, _func=1, output_level=0, compliance=0.1):
        self.__configure_source(_func)
        self.__configure_compliance(compliance)
        self.set_output(output_level)

    def enable_output(self, out=False):
        if out is True:
            self.inst.write("smua.source.output = 1")
            return
        self.inst.write("smua.source.output = 0")

    def get_current(self):
        return float(self.inst.query("printnumber(smua.measure.i())").split("\n")[0])
    def get_voltage(self):
        """
        Query the device for a voltage reading
        :return: float representation of the current measured
        """

        return float(self.inst.query("printnumber(smua.measure.v())").split("\n")[0])
    def powerDownPSU(self, speed=1):
        voltage=self.get_voltage()
        timeStep=.5 #seccond delay per step
        voltStep=speed*timeStep
        totalSteps=int(abs(voltage/voltStep))
        print("Powering down Keithley from %s to 0V."%voltage)
        print("At a rate of %sV/s with delay of %s, there will be %s steps at %sV per step for %s secconds."%(speed,timeStep,totalSteps,voltStep,totalSteps*timeStep))
        voltages=linspace(voltage,0,totalSteps)
        for volt in voltages:
            self.set_output(volt)
            time.sleep(timeStep)
