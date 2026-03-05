import serial
import time
import serial.tools.list_ports
import re

def available_com_ports():
    comport = []
    for port in serial.tools.list_ports.comports():
        comport.append(port.name)
    comport = list(comport)
    return list(comport)

class odl_connect():
    def __init__(self,port_name, odl_type):
        self.command_terminate = '\r\n'
        self.port_name = port_name
        self.serial_device = None
        self.odl_type: str = odl_type
        self.select_com_ports(port_name)

    def select_com_ports(self, ports):
        try:
            self.serial_device = serial.Serial(ports, baudrate=9600)
        except Exception as e:
            exit()
    
    def get_serial(self):
        cmd = 'V2'
        response = self.serial_command(cmd)
        if self.odl_type.endswith('167PS'):
            return response.replace(cmd, "").replace("Done", "").replace("\r", "").replace("\n", "")
        else:
            # return response.split('Done')[0].split('\r\n')[1]
            return response


    def get_device_info(self):
        cmd = 'V1'
        response = self.serial_command(cmd)

        if self.odl_type.endswith('167PS'):
            response = response.replace(cmd, "").replace("\r", "").replace("\n", "").replace("Done", "")
            device_name = response.split(" ")[0]
            hwd_version = response.split(" ")[1]
        else:
            return response

        return device_name, hwd_version

    def get_mfg_date(self):
        cmd = 'd?'
        response = self.serial_command(cmd)

        if self.odl_type.endswith('167PS'):
            response = response.replace(cmd, "").replace("\r", "").replace("\n", "").replace("Done", "")
            date = response
        else:
            date = response

        return date

    def echo(self, on_off):
        cmd = 'e' + str(on_off)
        response = self.serial_command(cmd)
        if self.odl_type.endswith('167PS'):
            return response.replace(cmd, "").replace("\r", "").replace("\n", "").replace("Done", "")
        else:
            return response
    
    def reset(self):
        cmd = 'RESET'
        response = self.serial_command(cmd)
        if self.odl_type.endswith('167PS'):
            return response.replace(cmd, "").replace("\r", "").replace("\n", "").replace("Done", "")
        else:
            return response
    
    def oz_mode(self, on_off):                  # on_off -> 0: OZ mode OFF | 1: OZ mode ON
        cmd = 'OZ-SHS' + str(on_off)
        response = self.serial_command(cmd)
        if self.odl_type.endswith('167PS'):
            return response.replace(cmd, "")
        else:
            return response
    
    def home(self):
        cmd = 'FH'
        response = self.serial_command(cmd, retries=450)
        if self.odl_type.endswith('167PS'):
            return response.replace(cmd, "").replace("\r", "").replace("\n", "")
        else:
            return response
    
    def forward(self):
        cmd = 'GF'
        response = self.serial_command(cmd, retries = 15)
        return response

    def reverse(self):
        cmd = 'GR'
        response = self.serial_command(cmd, retries = 15)
        return response

    def stop(self):
        cmd = 'G0'
        response = self.serial_command(cmd)
        return response

    def set_step(self, value):
        cmd = 'S' + str(value)
        response = self.serial_command(cmd)
        return response
        
    def get_step(self):
        cmd = 'S?'
        response = self.serial_command(cmd)
        if self.odl_type.endswith("ODL-300"):
            step = response.replace("\r", "").replace("\n", "").replace("Done", "").replace("STEP:", "")
        else:
            step = response.split('Done')[0].split(':')[-1]

    # Keep only digits
        step = re.sub(r"\D", "", step)

        return int(step) if step else 0

    def write_to_flash(self):
        cmd = 'OW'
        if self.odl_type.endswith("ODL-300"):
            response = 'Done'
        else:
            response = self.serial_command(cmd)
        return response
    
    def start_burn_in(self, parameter):
        cmd = 'OZBI' + str(parameter)
        response = self.serial_command(cmd)
        # for ODL 167PS this command does not exist, so assume success response
        if self.odl_type.endswith('167PS'):
            response = 'Done'
        return response

    def write_name(self,parameter):
        if self.odl_type.endswith('167PS'):
            cmd = f"OD,{parameter},{self.get_serial()},{self.get_mfg_date()}"
            response = self.serial_command(cmd)
        else:
            cmd = 'ODN' + str(parameter)
            response = self.serial_command(cmd)

        return response
    
    def write_serial(self,parameter):
        if self.odl_type.endswith("167PS"):
            cmd = f"OD,{self.get_device_info()[0]},{parameter},{self.get_mfg_date()}"
            response = self.serial_command(cmd)
        else:
            cmd = 'ODS' + str(parameter)
            response = self.serial_command(cmd)
        return response
    
    def write_mfg_date(self,parameter):
        if self.odl_type.endswith("167PS"):
            cmd = f"OD,{self.get_device_info()[0]},{self.get_serial()},{parameter}"
            response = self.serial_command(cmd)
        else:
            cmd = 'ODM' + str(parameter)

        response = self.serial_command(cmd)
        return response

    def write_hw_version(self,parameter):
        if self.odl_type.endswith("167PS"):
            cmd = f"OHW,{parameter}"
        else:
            cmd = 'OHW' + str(parameter)
        response = self.serial_command(cmd)
        return response

    def serial_close(self):
        self.serial_device.close()

    def serial_send(self,serial_cmd ):
        # Encode and send the command to the serial device.
        self.serial_device.flushInput()       #flush input buffer, discarding all its contents
        self.serial_device.flushOutput()      #flush output buffer, aborting current output and discard all that is in buffer        
        self.serial_device.write(serial_cmd.encode())

    def serial_read(self, retries=10):
        # The Python serial "in_waiting" property is the count of bytes available
        # for reading at the serial port.  If the value is greater than zero
        # then we know we have content available.
        device_output = ''
        got_OK = False
        while self.serial_device.in_waiting > 0 or (got_OK is False and retries > 0):
            device_output += \
                (self.serial_device.read(self.serial_device.in_waiting)).decode('iso-8859-1')
            # command output is complete.
            if device_output.find("Done") >= 0:
                got_OK = True
            if self.odl_type.endswith('167PS'):
                time.sleep(0.1)
            else:
                time.sleep(0.05)
            retries -= 1
        return device_output

    def serial_command(self, serial_cmd, retries=5):
        self.serial_send(serial_cmd + self.command_terminate)
        device_output = self.serial_read( retries)
        return device_output

    def readKey(self, key, retries = 5): 
        device_output = ''
        got_OK = False
        while self.serial_device.in_waiting > 0 or (got_OK is False and retries > 0):            
            device_output += \
                (self.serial_device.read(self.serial_device.in_waiting)).decode('iso-8859-1')
            # command output is complete.
            if device_output.find(key) >= 0:
                got_OK = True
            time.sleep(0.05)
            retries -= 1
        return device_output        

    def readall(self, sectimeout=5): 
            ok = False
            bytes = self.serial_device.read(1)
            while self.serial_device.inWaiting() > 0:
                bytes += self.serial_device.read(1)
            msg = bytes.decode('UTF-8')
            ok = True
            return ok, msg


    