import csv
import wx
from motor_controller import odl_connect, available_com_ports

class EventSink(object):
    def __init__(self, frame):
        self.counter = 0
        self.frame = frame
    def DataReady(self):
        self.counter += 1
        self.frame.SetTitle("DataReady fired {0} times".format(self.counter))

class AppLogic:
    def __init__(self, gui):
        self.gui = gui
        self.motor = None

    # --- DataRay Logic ---
    def write_to_csv(self):
        rb_selection = self.gui.rb.GetStringSelection()
        try:
            if rb_selection == "WinCam":
                data = [[x] for x in self.gui.gd.ctrl.GetWinCamDataAsVariant()]
            else:
                p_sel = self.gui.cb.GetStringSelection()
                if p_sel == "Profile_X":
                    data = [[x] for x in self.gui.px.ctrl.GetProfileDataAsVariant()]
                elif p_sel == "Profile_Y":
                    data = [[x] for x in self.gui.py.ctrl.GetProfileDataAsVariant()]
                else:
                    dx = self.gui.px.ctrl.GetProfileDataAsVariant()
                    dy = self.gui.py.ctrl.GetPRofileDataAsVariant()
                    data = [list(row) for row in zip(dx, dy)]
            
            with open(self.gui.ti.GetValue(), 'w') as fp:
                csv.writer(fp, delimiter=',').writerows(data)
        except Exception as e: print(f"CSV Error: {e}")

    def process_xc_timer(self):
        try:
            xc = self.gui.gd.ctrl.GetOcxResult(82)
            print("Wva:", xc)
        except: pass

    # --- Motor Logic ---
    def connect_motor(self):
        port = self.gui.port_dropdown.GetValue()
        m_type = self.gui.type_dropdown.GetValue()
        if port:
            self.motor = odl_connect(port, m_type)
            print(f"Connected to {m_type} on {port}")

    def motor_command(self, action, value=None):
        if not self.motor: return
        if action == "home": self.motor.home()
        elif action == "forward": self.motor.forward()
        elif action == "reverse": self.motor.reverse()
        elif action == "stop": self.motor.stop()
        elif action == "set_step": self.motor.set_step(value)
