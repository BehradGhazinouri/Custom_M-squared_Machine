import wx
import wx.lib.activex
import comtypes.client
from logic import AppLogic, EventSink, available_com_ports

class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title='DataRay & Motor Control', size=(1250, 950))
        self.logic = AppLogic(self)
        p = wx.Panel(self)
        
        # Main Layout: Left for Buttons, Right for Diagrams
        main_layout = wx.BoxSizer(wx.HORIZONTAL)
        
        # --- LEFT PANEL: ALL CONTROLS ---
        left_panel = wx.Panel(p, size=(300, -1))
        
        # DataRay Setup (Hidden Control)
        self.gd = wx.lib.activex.ActiveXCtrl(left_panel, 'DATARAYOCX.GetDataCtrl.1')
        sink_obj = EventSink(self)
        self.sink = comtypes.client.GetEvents(self.gd.ctrl, sink_obj)
        
        # DataRay Buttons
        self._add_ax_button(left_panel, 297, (260, 50), (10, 10))   # Main Start/Stop
        self._add_ax_button(left_panel, 171, (125, 25), (10, 65))
        self._add_ax_button(left_panel, 172, (125, 25), (145, 65))

        # File/Save Section
        wx.StaticText(left_panel, label="File:", pos=(10, 105))
        self.ti = wx.TextCtrl(left_panel, value="C:/Users/bghazinouri/Documents/DA station/Organized_rev3/sweep.csv", pos=(40, 105), size=(220, -1))
        self.rb = wx.RadioBox(left_panel, label="Data Source:", pos=(10, 135), choices=["Profile", "WinCam"])
        self.cb = wx.ComboBox(left_panel, pos=(10, 195), choices=["Profile_X", "Profile_Y", "Both"], size=(250, -1))
        self.cb.SetSelection(0)
        btn_write = wx.Button(left_panel, label="Write CSV Manually", pos=(10, 225), size=(250, -1))
        btn_write.Bind(wx.EVT_BUTTON, lambda e: self.logic.write_to_csv())

        # Motor Connection
        wx.StaticLine(left_panel, pos=(10, 265), size=(260, 2))
        wx.StaticText(left_panel, label="Motor Port:", pos=(10, 280))
        self.port_dropdown = wx.ComboBox(left_panel, pos=(90, 280), choices=available_com_ports())
        wx.StaticText(left_panel, label="Type:", pos=(10, 310))
        self.type_dropdown = wx.ComboBox(left_panel, pos=(90, 310), choices=["167PS", "ODL-300"], value="167PS")
        btn_conn = wx.Button(left_panel, label="Connect Motor", pos=(10, 340), size=(130, -1))
        btn_conn.Bind(wx.EVT_BUTTON, lambda e: self.logic.connect_motor())

        # Position Display
        wx.StaticText(left_panel, label="Pos:", pos=(155, 345))
        self.pos_display = wx.StaticText(left_panel, label="0.000", pos=(190, 345))
        self.pos_display.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        self.pos_display.SetForegroundColour(wx.BLUE)

        # Movement Buttons
        btn_home = wx.Button(left_panel, label="Home", pos=(10, 375), size=(80, -1))
        btn_home.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("home"))
        btn_fwd = wx.Button(left_panel, label="Forward", pos=(95, 375), size=(80, -1))
        btn_fwd.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("forward"))
        btn_rev = wx.Button(left_panel, label="Reverse", pos=(180, 375), size=(80, -1))
        btn_rev.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("reverse"))

        btn_stop = wx.Button(left_panel, label="STOP MOTOR", pos=(10, 410), size=(250, 40))
        btn_stop.SetBackgroundColour('red')
        btn_stop.Bind(wx.EVT_BUTTON, lambda e: self.logic.motor_command("stop"))

        # --- SWEEP FUNCTION UI (Added) ---
        wx.StaticLine(left_panel, pos=(10, 465), size=(260, 2))
        wx.StaticText(left_panel, label="Sweep Start (mm):", pos=(10, 480))
        self.sweep_start = wx.TextCtrl(left_panel, value="0.0", pos=(140, 480), size=(100, -1))
        
        wx.StaticText(left_panel, label="Increment (mm):", pos=(10, 510))
        self.sweep_inc = wx.TextCtrl(left_panel, value="0.5", pos=(140, 510), size=(100, -1))
        
        wx.StaticText(left_panel, label="Num Increments:", pos=(10, 540))
        self.sweep_count = wx.TextCtrl(left_panel, value="5", pos=(140, 540), size=(100, -1))
        
        btn_sweep = wx.Button(left_panel, label="RUN AUTOMATED SWEEP", pos=(10, 575), size=(250, 40))
        btn_sweep.SetBackgroundColour(wx.Colour(144, 238, 144)) # Light Green
        btn_sweep.Bind(wx.EVT_BUTTON, lambda e: self.logic.run_sweep())

        # --- RIGHT PANEL: ALL DIAGRAMS ---
        right_panel = wx.Panel(p)
        
        # 3D View and CCD Image (Top Row)
        wx.lib.activex.ActiveXCtrl(right_panel, size=(450, 400), axID='DATARAYOCX.ThreeDviewCtrl.1', pos=(10, 10))
        wx.lib.activex.ActiveXCtrl(right_panel, size=(450, 400), axID='DATARAYOCX.CCDimageCtrl.1', pos=(470, 10))
        
        # Profiles (Bottom Row)
        self.px = wx.lib.activex.ActiveXCtrl(right_panel, size=(450, 350), axID='DATARAYOCX.ProfilesCtrl.1', pos=(10, 420))
        self.px.ctrl.ProfileID = 22
        
        self.py = wx.lib.activex.ActiveXCtrl(right_panel, size=(450, 350), axID='DATARAYOCX.ProfilesCtrl.1', pos=(470, 420))
        self.py.ctrl.ProfileID = 23

        # Final Assembly
        main_layout.Add(left_panel, 0, wx.EXPAND | wx.ALL, 10)
        main_layout.Add(right_panel, 1, wx.EXPAND | wx.ALL, 10)
        p.SetSizer(main_layout)

        # Timer setup
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, lambda e: self.logic.process_xc_timer(), self.timer)
        self.gd.ctrl.StartDriver()
        self.timer.Start(500) 

    def _add_ax_button(self, parent, bid, size, pos):
        btn = wx.lib.activex.ActiveXCtrl(parent=parent, size=size, pos=pos, axID='DATARAYOCX.ButtonCtrl.1')
        btn.ctrl.ButtonID = bid
        return btn
