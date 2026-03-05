import csv
import wx
from motor_controller import odl_connect, available_com_ports

# Conversion Factor
STEPS_PER_MM = 1390

class EventSink(object):
    def __init__(self, frame):
        self.counter = 0
        self.frame = frame
    def DataReady(self):
        self.counter += 1
        self.frame.SetTitle(f"DataReady fired {self.counter} times")

class AppLogic:
    def __init__(self, gui):
        self.gui = gui
        self.motor = None

    def connect_motor(self):
        port = self.gui.port_dropdown.GetValue()
        m_type = self.gui.type_dropdown.GetValue()
        if port:
            self.motor = odl_connect(port, m_type)
            print(f"Connected to {m_type} on {port}")

    def process_xc_timer(self):
        try:
            # DataRay result polling
            xc = self.gui.gd.ctrl.GetOcxResult(82)
        except: pass
        
        # Real-time Position Conversion
        if self.motor:
            try:
                raw_steps = self.motor.get_step()
                # Math: steps / conversion_factor = mm
                mm_pos = raw_steps / STEPS_PER_MM
                self.gui.pos_display.SetLabel(f"{mm_pos:.3f}")
            except Exception as e:
                print(f"Update error: {e}")

    def motor_command(self, action, value=None):
        if not self.motor: return
        if action == "home": self.motor.home()
        elif action == "forward": self.motor.forward()
        elif action == "reverse": self.motor.reverse()
        elif action == "stop": self.motor.stop()

    def write_to_csv(self):
        # Existing CSV logic
        pass
