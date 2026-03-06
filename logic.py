import csv
import wx
import threading
import time
import queue
import os
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
        self.sweep_active = False
        self.capture_queue = queue.Queue()

    def connect_motor(self):
        port = self.gui.port_dropdown.GetValue()
        m_type = self.gui.type_dropdown.GetValue()
        if port:
            self.motor = odl_connect(port, m_type)
            print(f"Connected to {m_type} on {port}")

    def process_xc_timer(self):
        # Only update pos if sweep isn't hogging the motor/ActiveX
        if self.motor and not self.sweep_active:
            try:
                raw_steps = self.motor.get_step()
                mm_pos = raw_steps / STEPS_PER_MM
                self.gui.pos_display.SetLabel(f"{mm_pos:.3f}")
            except: pass

    def motor_command(self, action, value=None):
        if not self.motor: return
        if action == "home": self.motor.home()
        elif action == "forward": self.motor.forward()
        elif action == "reverse": self.motor.reverse()
        elif action == "stop": self.motor.stop()

    def run_sweep(self):
        if not self.motor:
            wx.MessageBox("Connect motor first!", "Error")
            return
        if self.sweep_active: return

        try:
            start = float(self.gui.sweep_start.GetValue())
            inc = float(self.gui.sweep_inc.GetValue())
            count = int(self.gui.sweep_count.GetValue())
            path = self.gui.ti.GetValue()

            self.sweep_active = True
            t = threading.Thread(target=self._sweep_worker, args=(start, inc, count, path))
            t.daemon = True
            t.start()
        except Exception as e:
            wx.MessageBox(f"Input Error: {e}")
            self.sweep_active = False

    def _grab_data(self):
        try:
            x = self.gui.gd.ctrl.GetOcxResult(78)
            y = self.gui.gd.ctrl.GetOcxResult(82)
            self.capture_queue.put((x, y))
        except:
            self.capture_queue.put((0, 0))

    def _sweep_worker(self, start, inc, count, path):
        data_log = []
        try:
            for i in range(count):
                target_mm = start + (i * inc)
                target_steps = int(target_mm * STEPS_PER_MM)
                
                # 1. Move to position
                self.motor.set_step(target_steps)
                
                # 2. VERIFY ARRIVAL: Poll motor until it is close to target steps
                arrived = False
                timeout_start = time.time()
                while not arrived and (time.time() - timeout_start < 10):
                    try:
                        current_steps = self.motor.get_step()
                        if abs(current_steps - target_steps) < 5: # Threshold of 5 steps
                            arrived = True
                    except: pass
                    time.sleep(0.5)
                
                # 3. MANDATORY DWELL: Wait in position for 0.5s
                time.sleep(0.5)
                
                # 4. RECORD: Capture data immediately after dwell
                wx.CallAfter(self._grab_data)
                
                # Pull data from the queue
                val_x, val_y = self.capture_queue.get(timeout=5)
                data_log.append([target_mm, val_x, val_y])
                print(f"Recorded Step {i+1}/{count}: {target_mm}mm")

            # 5. HARD SAVE TO DISK
            with open(path, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Distance (mm)", "X (78)", "Y (82)"])
                writer.writerows(data_log)
                f.flush()
                os.fsync(f.fileno()) # Force Windows to write the file now
            
            wx.CallAfter(lambda: wx.MessageBox(f"Success! {len(data_log)} steps saved to {path}", "Sweep Done"))
        
        except Exception as e:
            wx.CallAfter(lambda: wx.MessageBox(f"Sweep failed: {str(e)}", "Error"))
        
        finally:
            self.sweep_active = False

    def write_to_csv(self): pass
