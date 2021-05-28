import math
import re
import base64

def calculate_dom_value(dom_bytes, dom_type):
    """
    Calculate module DOM vale from raw eeprom
    :param list dom_bytes:list of 2-byte of DDM value in module's EEPROM
    :param str dom_type: type of DOM to calculate
     Accepted sting are 'temperature', 'supply_voltage', 'tx_bias_current', 'tx_power', and 'rx_power'
    :return: DOM value (temperature in Celsius, voltage in Volt, current in mA, power in dBm)
    :rtype:float
    """

    if dom_type == 'temperature':
        value = (twos_comp(int(256 * dom_bytes[0]) + dom_bytes[1])) / 256.0
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


def twos_comp(in_val, in_bits=16):
    """
    compute the 2's complement of int value in_val
    based on an integer length of in_bits bits.
    """
    if (in_val & (1 << (in_bits - 1))) != 0: # if sign bit is set e.g., 8bit: 128-255
        in_val = in_val - (1 << in_bits)        # compute negative value
    return in_val  # return positive value as is


def get_bit_value(byte_data, start_bit, num_bit):
    """ From byte data, return bit data

    :param str byte_data: data in hex format for 1 byte
    :param int start_bit: start bit location
    :param int num_bit: number of bit from start bit to return
    :return: bin value in binary format
    :rtype: str
    """
    value = "{0:08b}".format(int(byte_data, 16))
    stop_bit = start_bit + num_bit
    value = value[::-1]
    value = value[start_bit:stop_bit]
    value = value[::-1]
    return value


def calculate_ber(num_error, test_data_rate, test_time):
    """ Calculate BER from number of error count, data rate, and traffic test time

    :param int num_error: number of error count
    :pa_ram str test_data_rate: test data rate e.g. 10GE, 25GE, 40GE, 100GE, 10GFC, 4/8/16/32GFC
    :param int test_time: traffic test time in second
    :return: BER
    :rtype: float
    """
    if test_data_rate == '10GE':
        n_bits = 10.3125*(10**9)*test_time
    elif test_data_rate == '25GE':
        n_bits = 25.78125*(10**9)*test_time
    elif test_data_rate == '40GE':
        n_bits = 41.25*(10**9)*test_time
    elif test_data_rate == '100GE':
        n_bits = 103.125*(10**9)*test_time
    elif test_data_rate == '10GFC':
        n_bits = 10.51875*(10**9)*test_time
    elif test_data_rate == '4GFC':
        n_bits = 4.25*(10**9)*test_time
    elif test_data_rate == '8GFC':
        n_bits = 8.50*(10**9)*test_time
    elif test_data_rate == '16GFC':
        n_bits = 14.025*(10**9)*test_time
    elif test_data_rate == '32GFC':
        n_bits = 28.050*(10**9)*test_time
    else:
        n_bits = -1
    ber = float(num_error) / float(n_bits)
    return '{0:.2e}'.format(ber)


def convert_ico_to_base64(ico):
    open_icon = open(ico, "rb")  # qq.icon is the icon you want to put in
    b64str = base64.b64encode(open_icon.read())  # Read in base64 format
    open_icon.close()
    write_data = "img=%s" % b64str
    f = open("icon.py", "w+")  # Write the data read above into the img array of qq.py
    f.write(write_data)
    f.close()


if __name__ == '__main__':
    # ddm_type = ['temperature', 'supply_voltage', 'tx_bias_current', 'tx_power', 'rx_power']
    # ddm_bytes = [int('0x' + '41', 16), int('0x' + '41', 16)]
    #
    # ddm_value = calculate_dom_value(dom_bytes=ddm_bytes, dom_type=ddm_type[3])
    # print(ddm_value)
    ico_name = "calculator.ico"
    convert_ico_to_base64(ico_name)

