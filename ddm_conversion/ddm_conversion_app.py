import tkinter as tk
import math

# init parameters
ddm_type = ["temperature", "supply_voltage", "tx_bias_current", "tx_power", "rx_power"]
data_rate = ["10GE", "25GE", "100GE", "10GFC", "4GFC", "8GFC", "16GFC", "32GFC"]

ddm_temp = []
ddm_volt = []
ddm_tx_bias = []
ddm_tx_power = []
ddm_rx_power = []

# init event
return_event = "<Return>"


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DDM Conversion - Tkinter")
        self.minsize(width=100, height=100)
        self.resizable(0, 0)

        # define widgets
        # DDM Temperature
        self.lbl_ddm_temp = tk.Label(self, text="DDM Temperature (0x):")
        self.ddm_temp_msb = tk.Entry(self, width=10)
        # self.ddm_temp_msb.bind(return_event, self.ddm_temp_changed)
        self.ddm_temp_msb.bind(return_event, lambda event, ddm=ddm_type[0]: self.ddm_changed(event, ddm_type[0]))
        self.ddm_temp_lsb = tk.Entry(self, width=10)
        # self.ddm_temp_lsb.bind(return_event, self.ddm_temp_changed)
        self.ddm_temp_lsb.bind(return_event, lambda event, ddm=ddm_type[0]: self.ddm_changed(event, ddm_type[0]))
        self.lbl1 = tk.Label(self, text="==>>")
        self.ddm_temp_value = tk.Text(self, height=1, width=8)
        self.unit1 = tk.Label(self, text="Degree C")

        # DDM Voltage
        self.lbl_ddm_volt = tk.Label(self, text="DDM Voltage (0x):")
        self.ddm_volt_msb = tk.Entry(self, width=10)
        self.ddm_volt_msb.bind(return_event, lambda event, ddm=ddm_type[1]: self.ddm_changed(event, ddm_type[1]))
        self.ddm_volt_lsb = tk.Entry(self, width=10)
        self.ddm_volt_lsb.bind(return_event, lambda event, ddm=ddm_type[1]: self.ddm_changed(event, ddm_type[1]))
        self.lbl2 = tk.Label(self, text="==>>")
        self.ddm_volt_value = tk.Text(self, height=1, width=8)
        self.unit2 = tk.Label(self, text="Volt")

        # DDM Tx Bias Current
        self.lbl_ddm_bias = tk.Label(self, text="DDM Bias Current (0x):")
        self.ddm_bias_msb = tk.Entry(self, width=10)
        self.ddm_bias_msb.bind(return_event, lambda event, ddm=ddm_type[2]: self.ddm_changed(event, ddm_type[2]))
        self.ddm_bias_lsb = tk.Entry(self, width=10)
        self.ddm_bias_lsb.bind(return_event, lambda event, ddm=ddm_type[2]: self.ddm_changed(event, ddm_type[2]))
        self.lbl3 = tk.Label(self, text="==>>")
        self.ddm_bias_value = tk.Text(self, height=1, width=8)
        self.unit3 = tk.Label(self, text="mA")

        # DDM Tx Power
        self.lbl_ddm_tx = tk.Label(self, text="DDM Tx Power (0x):")
        self.ddm_tx_msb = tk.Entry(self, width=10)
        self.ddm_tx_msb.bind(return_event, lambda event, ddm=ddm_type[3]: self.ddm_changed(event, ddm_type[3]))
        self.ddm_tx_lsb = tk.Entry(self, width=10)
        self.ddm_tx_lsb.bind(return_event, lambda event, ddm=ddm_type[3]: self.ddm_changed(event, ddm_type[3]))
        self.lbl4 = tk.Label(self, text="==>>")
        self.ddm_tx_value = tk.Text(self, height=1, width=8)
        self.unit4 = tk.Label(self, text="dBm")

        # DDM Rx Power
        self.lbl_ddm_rx = tk.Label(self, text="DDM Rx Power (0x):")
        self.ddm_rx_msb = tk.Entry(self, width=10)
        self.ddm_rx_msb.bind(return_event, lambda event, ddm=ddm_type[4]: self.ddm_changed(event, ddm_type[4]))
        self.ddm_rx_lsb = tk.Entry(self, width=10)
        self.ddm_rx_lsb.bind(return_event, lambda event, ddm=ddm_type[4]: self.ddm_changed(event, ddm_type[4]))
        self.lbl5 = tk.Label(self, text="==>>")
        self.ddm_rx_value = tk.Text(self, height=1, width=8)
        self.unit5 = tk.Label(self, text="dBm")

        # # BER
        # self.lbl_ber_conv = tk.Label(self, text="BER-CONVERSION:")
        # self.lbl_data_rate = tk.Label(self, text="Data Rate")
        #
        # variable = tk.StringVar(self)
        # variable.set(data_rate[0])    # set default value
        # self.data_rate_dropdown = tk.OptionMenu(self, variable, *data_rate)
        #
        # self.lbl_err_num = tk.Label(self, text="Err num (0x)")
        # self.err_num = tk.Entry(self, width=12)
        # self.lbl_trf_time = tk.Label(self, text="Test Time")
        # self.trf_time = tk.Entry(self, width=10)
        # self.lbl_ber = tk.Label(self, text="BER = ")
        # self.ber_value = tk.Text(self, height=1, width=8)

        # widgets position
        # DDM Temperature
        self.lbl_ddm_temp.grid(row=0, column=1)
        self.ddm_temp_msb.grid(row=0, column=2, pady=5)
        self.ddm_temp_lsb.grid(row=0, column=3, padx=3, pady=5)
        self.lbl1.grid(row=0, column=4, pady=5)
        self.ddm_temp_value.grid(row=0, column=5, pady=5)
        self.unit1.grid(row=0, column=6, pady=5)

        # DDM Voltage
        self.lbl_ddm_volt.grid(row=2, column=1, pady=5)
        self.ddm_volt_msb.grid(row=2, column=2, pady=5)
        self.ddm_volt_lsb.grid(row=2, column=3, padx=3, pady=5)
        self.lbl2.grid(row=2, column=4, pady=5)
        self.ddm_volt_value.grid(row=2, column=5, pady=5)
        self.unit2.grid(row=2, column=6, pady=5)

        # DDM Tx Bias Current
        self.lbl_ddm_bias.grid(row=3, column=1, pady=5)
        self.ddm_bias_msb.grid(row=3, column=2, pady=5)
        self.ddm_bias_lsb.grid(row=3, column=3, padx=3, pady=5)
        self.lbl3.grid(row=3, column=4, pady=5)
        self.ddm_bias_value.grid(row=3, column=5, pady=5)
        self.unit3.grid(row=3, column=6, pady=5)

        # DDM Tx Power
        self.lbl_ddm_tx.grid(row=4, column=1, pady=5)
        self.ddm_tx_msb.grid(row=4, column=2, pady=5)
        self.ddm_tx_lsb.grid(row=4, column=3, padx=3, pady=5)
        self.lbl4 .grid(row=4, column=4, pady=5)
        self.ddm_tx_value.grid(row=4, column=5, pady=5)
        self.unit4.grid(row=4, column=6, pady=5)

        # DDM Rx Power
        self.lbl_ddm_rx.grid(row=5, column=1, pady=5)
        self.ddm_rx_msb.grid(row=5, column=2, pady=5)
        self.ddm_rx_lsb.grid(row=5, column=3, padx=3, pady=5)
        self.lbl5.grid(row=5, column=4, pady=5)
        self.ddm_rx_value.grid(row=5, column=5, pady=5)
        self.unit5.grid(row=5, column=6, pady=5)

        # # BER
        # self.lbl_ber_conv.grid(row=6, column=1, pady=5)
        # self.lbl_data_rate.grid(row=7, column=1, padx=5)
        # self.data_rate_dropdown.grid(row=8, column=1, pady=5)
        # self.lbl_err_num.grid(row=8, column=2, padx=5)
        # self.err_num.grid(row=8, column=3, padx=5)
        # self.lbl_trf_time.grid(row=8, column=4, padx=5)
        # self.trf_time.grid(row=8, column=5, padx=5)

    def ddm_changed(self, event, ddm):
        global ddm_temp, ddm_volt, ddm_tx_bias, ddm_tx_power, ddm_rx_power
        ddm_temp = []
        ddm_volt = []
        ddm_tx_bias = []
        ddm_tx_power = []
        ddm_rx_power = []

        if ddm_type[0] == ddm:
            print("DDM Temperature")
            msb = self.ddm_temp_msb.get()
            lsb = self.ddm_temp_lsb.get()
            if msb == "":
                self.ddm_temp_msb.insert(0, "00")
                ddm_temp.append(int("0x" + "00", 16))
            else:
                ddm_temp.append(int("0x" + msb, 16))
            if lsb == "":
                self.ddm_temp_lsb.insert(0, "00")
                ddm_temp.append(int("0x" + "00", 16))
            else:
                ddm_temp.append(int("0x" + lsb, 16))
            result = self.calculate_dom_value(dom_bytes=ddm_temp, dom_type=ddm_type[0])
            print(result)
            self.ddm_temp_value.delete("1.0", "end")
            self.ddm_temp_value.insert("insert", result)

        if ddm_type[1] == ddm:
            print("DDM Voltage")
            msb = self.ddm_volt_msb.get()
            lsb = self.ddm_volt_lsb.get()
            if msb == "":
                self.ddm_volt_msb.insert(0, "00")
                ddm_volt.append(int("0x" + "00", 16))
            else:
                ddm_volt.append(int("0x" + msb, 16))
            if lsb == "":
                self.ddm_volt_lsb.insert(0, "00")
                ddm_volt.append(int("0x" + "00", 16))
            else:
                ddm_volt.append(int("0x" + lsb, 16))
            result = self.calculate_dom_value(dom_bytes=ddm_volt, dom_type=ddm_type[1])
            print(result)
            self.ddm_volt_value.delete("1.0", "end")
            self.ddm_volt_value.insert("insert", result)

        if ddm_type[2] == ddm:
            print("DDM Tx Bias Current")
            msb = self.ddm_bias_msb.get()
            lsb = self.ddm_bias_lsb.get()
            if msb == "":
                self.ddm_bias_msb.insert(0, "00")
                ddm_tx_bias.append(int("0x" + "00", 16))
            else:
                ddm_tx_bias.append(int("0x" + msb, 16))
            if lsb == "":
                self.ddm_bias_lsb.insert(0, "00")
                ddm_tx_bias.append(int("0x" + "00", 16))
            else:
                ddm_tx_bias.append(int("0x" + lsb, 16))
            result = self.calculate_dom_value(dom_bytes=ddm_tx_bias, dom_type=ddm_type[2])
            print(result)
            self.ddm_bias_value.delete("1.0", "end")
            self.ddm_bias_value.insert("insert", result)

        if ddm_type[3] == ddm:
            print("DDM Tx Power")
            msb = self.ddm_tx_msb.get()
            lsb = self.ddm_tx_lsb.get()
            if msb == "":
                self.ddm_tx_msb.insert(0, "00")
                ddm_tx_power.append(int("0x" + "00", 16))
            else:
                ddm_tx_power.append(int("0x" + msb, 16))
            if lsb == "":
                self.ddm_bias_lsb.insert(0, "00")
                ddm_tx_power.append(int("0x" + "00", 16))
            else:
                ddm_tx_power.append(int("0x" + lsb, 16))
            result = self.calculate_dom_value(dom_bytes=ddm_tx_power, dom_type=ddm_type[3])
            print(result)
            self.ddm_tx_value.delete("1.0", "end")
            self.ddm_tx_value.insert("insert", result)

        if ddm_type[4] == ddm:
            print("DDM Rx Power")
            msb = self.ddm_rx_msb.get()
            lsb = self.ddm_rx_lsb.get()
            if msb == "":
                self.ddm_rx_msb.insert(0, "00")
                ddm_rx_power.append(int("0x" + "00", 16))
            else:
                ddm_rx_power.append(int("0x" + msb, 16))
            if lsb == "":
                self.ddm_rx_lsb.insert(0, "00")
                ddm_rx_power.append(int("0x" + "00", 16))
            else:
                ddm_rx_power.append(int("0x" + lsb, 16))
            result = self.calculate_dom_value(dom_bytes=ddm_rx_power, dom_type=ddm_type[4])
            print(result)
            self.ddm_rx_value.delete("1.0", "end")
            self.ddm_rx_value.insert("insert", result)

    def calculate_dom_value(self, dom_bytes, dom_type):
        """
        Calculate module DOM vale from raw eeprom
        :param list dom_bytes:list of 2-byte of DDM value in module's EEPROM
        :param str dom_type: type of DOM to calculate
         Accepted sting are 'temperature', 'supply_voltage', 'tx_bias_current', 'tx_power', and 'rx_power'
        :return: DOM value (temperature in Celsius, voltage in Volt, current in mA, power in dBm)
        :rtype:float
        """

        if dom_type == 'temperature':
            value = (self.twos_comp(int(256 * dom_bytes[0]) + dom_bytes[1])) / 256.0
        elif dom_type == 'supply_voltage':
            value = (int(256 * dom_bytes[0]) + dom_bytes[1]) / 10000.0
        elif dom_type == 'tx_bias_current':
            value = 0.002 * (int(256 * dom_bytes[0]) + dom_bytes[1])
        elif dom_type == 'tx_power' or dom_type == 'rx_power':
            if dom_bytes[0] == 0 and dom_bytes[1] == 0:
                value = -999.9
            else:
                value = 10.0 * math.log10((256 * dom_bytes[0] + dom_bytes[1]) / 10000.0)
        else:
            value = -999.9
        return round(value, 2)

    def twos_comp(self, in_val, in_bits=16):
        """
        compute the 2's complement of int value in_val
        based on an integer length of in_bits bits.
        """
        if (in_val & (1 << (in_bits - 1))) != 0: # if sign bit is set e.g., 8bit: 128-255
            in_val = in_val - (1 << in_bits)        # compute negative value
        return in_val  # return positive value as is

    def calculate_ber(self, num_error, test_data_rate, test_time):
        """ Calculate BER from number of error count, data rate, and traffic test time

        :param int num_error: number of error count
        :pa_ram str test_data_rate: test data rate e.g. 10GE, 25GE, 40GE, 100GE, 10GFC, 4/8/16/32GFC
        :param int test_time: traffic test time in second
        :return: BER
        :rtype: float
        """
        if test_data_rate == '10GE':
            n_bits = 10.3125 * (10 ** 9) * test_time
        elif test_data_rate == '25GE':
            n_bits = 25.78125 * (10 ** 9) * test_time
        elif test_data_rate == '40GE':
            n_bits = 41.25 * (10 ** 9) * test_time
        elif test_data_rate == '100GE':
            n_bits = 103.125 * (10 ** 9) * test_time
        elif test_data_rate == '10GFC':
            n_bits = 10.51875 * (10 ** 9) * test_time
        elif test_data_rate == '4GFC':
            n_bits = 4.25 * (10 ** 9) * test_time
        elif test_data_rate == '8GFC':
            n_bits = 8.50 * (10 ** 9) * test_time
        elif test_data_rate == '16GFC':
            n_bits = 14.025 * (10 ** 9) * test_time
        elif test_data_rate == '32GFC':
            n_bits = 28.050 * (10 ** 9) * test_time
        else:
            n_bits = -1
        ber = float(num_error) / float(n_bits)
        return '{0:.2e}'.format(ber)


app = Application()
app.mainloop()
