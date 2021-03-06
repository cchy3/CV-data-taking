import visa

#This is used to connect to devices using gpib
class Instrument:
    def __init__(self):
        self.inst=None

    def getName(self,inst=None):
        if inst is None:
            if self.inst is not None:
                inst=self.inst
            else:
                raise(Exception("Called getName without an Instrument selected."))
        return str(inst.query('*IDN?')).lower()


    def test(self):
        return "working"

    #Args is a list of args
    def connect(self,*args):
        rm=visa.ResourceManager()
        for device in rm.list_resources():
            self.inst=rm.open_resource(device)
            idn=self.getName(self.inst).lower()
            for arg in args:
                if not arg.lower() in idn:
                    self.inst=None
                    break
            if self.inst is not None: break;
        if self.inst is None:
            raise Exception("The device %s was not found."%' '.join(args))
        else:
            print("Connected to %s"%' '.join(args))
            return self.inst

    def reset(self):
        self.inst.write("*RST;")
