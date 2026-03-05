import wx
import wx.lib.activex
import comtypes.client
from logic import AppLogic, EventSink, available_com_ports

class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title='DataRay & Motor Control', size=(1200, 900))
        self.logic = AppLogic(self)
        p = wx.Panel(self)

        # --- DataRay Section (ActiveX Core) ---
        self.gd = wx.lib.activex.ActiveXCtrl(p, 'DATARAYOCX.GetDataCtrl.1')
        
        # EventSink Setup
        sink_obj = EventSink(self)
        self.sink = comtypes.client.GetEvents(self.gd.ctrl, sink_obj)
        
        # --- Camera Control Buttons (Restored) ---
        self._add_ax_button(p, 297, (200, 50), (7, 0))   # Main Start/Stop
        self._add_ax_button(p, 171, (100, 25), (5, 55))  # Control Button 1
        self._add_ax_button(p, 172, (100, 25), (110, 55))# Control Button 2
        self._add_ax_button(p, 177, (100, 25), (5, 85))  # Control Button 3
        self._add_ax_button(p, 179, (100, 25), (110, 85))# Control Button 4

        # --- Visualizations (ActiveX) ---
        # 3D View
        wx.lib.activex.ActiveXCtrl(p, size=(250, 250), axID='DATARAYOCX.ThreeDviewCtrl.1', pos=(500, 250))
        
        # Profile X
        self.px = wx.lib.activex.ActiveXCtrl(p, size=(300, 200), axID='DATARAYOCX.ProfilesCtrl.1', pos=(0, 600))
        self.px.ctrl.ProfileID = 22
        
        # Profile Y
        self.py = wx.lib.activex.ActiveXCtrl(p, size=(300, 200), axID='DATARAYOCX.ProfilesCtrl.1', pos=(600, 600))
        self.py.ctrl.ProfileID = 23
        
        # CCD Image
        wx.lib.activex.ActiveXCtrl(p, axID='DATARAYOCX.CCDimageCtrl.1', size=(250, 250), pos=(250, 250))

        # --- Data Saving Controls (Native wx) ---
        wx.StaticText(p, label="File:", pos=(5, 115))
        self.ti = wx.TextCtrl(p, value="C:/Users/Public/Documents/output.csv", pos=(30, 115), size=(170, -1))
        self.rb = wx.RadioBox(p, label="Data:", pos=(5, 140), choices=["Profile", "WinCam"])
        self.cb = wx.ComboBox(p, pos=(5, 200), choices=["Profile_X", "Profile_Y", "Both"])
        self.cb.SetSelection(0)
        
        btn_write = wx.Button(p, label="Write", pos=(5, 225))
        btn_write.Bind(wx.EVT_BUTTON, lambda e: self.logic.write_to_csv())

        # --- Motor Control Section ---
        wx.StaticText(p, label="Motor Port:", pos=(5, 300))
        self.port_dropdown = wx.ComboBox(p, pos=(80, 300), choices=available_com_ports())
        
        wx.StaticText(p, label="Type:", pos=(5, 330))
        self.type_dropdown = wx.ComboBox(p, pos=(80, 330), choices=["167PS", "ODL-300"], value="167PS")
        
        btn_conn = wx.Button(p, label="Connect Motor", pos=(5, 360))
        btn_conn.Bind(wx.EVT_BUTTON, lambda e: self.logic.connect_motor())

        btn_home = wx.Button(p, label="Home", pos=(5, 390))
        btn_home.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("home"))

        btn_fwd = wx.Button(p, label="Forward", pos=(90, 390))
        btn_fwd.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("forward"))

        btn_rev = wx.Button(p, label="Reverse", pos=(175, 390))
        btn_rev.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("reverse"))

        btn_stop = wx.Button(p, label="STOP", pos=(5, 420), size=(255, 30))
        btn_stop.SetBackgroundColour('red')
        btn_stop.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("stop"))

        # --- Timer & Driver Init ---
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, lambda e: self.logic.process_xc_timer(), self.timer)
        
        self.gd.ctrl.StartDriver()
        self.timer.Start(500)

    def _add_ax_button(self, parent, bid, size, pos):
        """Helper to create ActiveX buttons with specific IDs"""
        btn = wx.lib.activex.ActiveXCtrl(parent=parent, size=size, pos=pos, axID='DATARAYOCX.ButtonCtrl.1')
        btn.ctrl.ButtonID = bid
        return btn
