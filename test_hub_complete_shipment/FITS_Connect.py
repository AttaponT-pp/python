from win32com import client

model = '*'
rev = '2.9.0.0'
fits_dll = client.Dispatch("FITSDLL.clsDB")

opn602_param = 'Packing Qty,RT,Build Type,PO No.,Part Number,MFG Part Number,' \
               'Supplier Name,Mother Lot Qty,Shipment Request Qty,Test Sampling type,SampQty'

opn1501_param = 'OPERATOR,Invoice No,Packing No,Packing Qty,RT,Build Type,PO No.,Part Number,MFG Part Number,' \
               'Supplier Name,Mother Lot Qty,Shipment Request Qty,Test Sampling type,SampQty,Pass Kitting Qty,' \
                'PID,Inspection Result'

opn702_param = 'OPERATOR,Invoice No,Serial No,Packing Qty,Packing No'


def init(opn):
    # Give location of dll
    # fits_dll = client.Dispatch("FITSDLL.clsDB")
    status = fits_dll.fn_InitDB(opn, rev, '')

    print("init=" + status)
    return status


def handshake(fits_dll, opn, inv):
    status = fits_dll.fn_handshake('*', opn, rev, inv)
    print("handshake = " + status)
    fits_dll.closeDB
    return status


def query(fits_dll, opn, sn, param, fs):
    # fn_query(model,operation,revision,serial,parameters[,fsp]);
    status = fits_dll.fn_query('*', opn, rev, sn,param,fs)
    print("query status: " + str(status))
    fits_dll.closeDB
    return status


def log(fits_dll ,opt, param, data, fs):
    # fn_log(model,operation,revision,parameters,values[,fsp]);
    global model, rev
    
    init = fits_dll.fn_InitDB(model, opt, rev, "")
    if init == 'False':
        return False
    status = fits_dll.fn_log(model,opt,rev,param,data,fs)
    print("log status: " + str(status))
    fits_dll.closeDB
    return status

def en_validate(opn, en):
    if not fits_dll.fn_InitDB(model, opn, rev, ""):
        return False
    status = fits_dll.fn_handshake('*', opn, rev, 'EN:' + en)
    fits_dll.closeDB
    if status == 'True':
        return {"status": True, "msg": ""}
    else:
        return {"status": False, "msg": status.split('|')[1]}
    
def valid_inv(opn, inv):
    if init(opn) == 'True':
        if handshake(fits_dll, opn, inv) == 'True':
            return {"status": True, "msg": ""}
        else:
            return {"status": False, "msg": "This invoice is not valid at {} operation".format(opn)}
    else:
        return {"status": False, "msg": "Cannot init FIT DB!"}


def get_necessory_data(fits_dll ,opn, rt, param):
    # fn_query(model,operation,revision,serial,parameters[,fsp]);
    # print 'input param: ' + param
    status = fits_dll.fn_query('*', opn, rev, rt, param, ',')
    output = status.split(',')
    fits_dll.closeDB
    return output


def record2fit(fits_dll, opn, param, data):
    if not fits_dll.fn_InitDB('*', '*', rev):
        return False
    status = fits_dll.fn_log('*', opn, rev, param, data, ',')
    print(status)
    fits_dll.closeDB
    return status


def get_sn_list(fits_dll, rt):
    if not fits_dll.fn_InitDB('*', '*', rev):
        return False
    sn_list = fits_dll.fn_query('*', '151', 'RT', '*', rt, ',')
    print(sn_list)
    fits_dll.closeDB
    return sn_list


def get_last_opn(fits_dll, sn):
    if not fits_dll.fn_InitDB('*', rev, ''):
        return False
    last_opn = fits_dll.fn_Query('*', "*", rev, sn, "last_opn", ',')
    fits_dll.closeDB
    return last_opn


# Check RTV Shipment Blocking Status
def check_block_rtv(fits_dll, etr):
    if init("*") == "True":
        # Get RT from opn.1303 ETR
        rt = fits_dll.fn_query("*", "1303", rev, etr, "RT", ',')
        # Check RT in Opn.924 RTV Shipment Blocking
        result = fits_dll.fn_query("*", "924", rev, rt, "RTV Shipment Blocking")
        print("RTV Shipment Blocking = {}".format(result))
        fits_dll.closeDB
        return result
    else:
        result = init("*")
        return result


def prepare_etr_info(fits_dll, etr):
    if fits_dll.fn_InitDB('*', rev, ''):
        # Get RT from opn.1303 ETR
        data1303 = fits_dll.fn_query("*", "1303", rev, etr, "Part Number,Supplier Name,RT,PO No.,Fail Qty", ',')
        print(data1303)
        rt = fits_dll.fn_query("*", "1303", rev, etr, "RT", ',')
        build_type = fits_dll.fn_query("*", "101", rev, rt, "Build Type", ',')
        # Check RT in Opn.924 RTV Shipment Blocking
        result = fits_dll.fn_query("*", "924", rev, rt, "RTV Shipment Blocking")
        print("RTV Shipment Blocking = {}".format(result))
        etr_data = data1303 + ',' + build_type + ',' + result
        print(etr_data)
        fits_dll.closeDB
        return etr_data
    else:
        result = fits_dll.fn_InitDB('*', rev, '')
        return result

def prepare_oba_info(fits_dll, packing_num):
    print("Input Packing no. = {}".format(packing_num))
    if not fits_dll.fn_InitDB('*', rev, ''):
        return False
    opn602_data = fits_dll.fn_query('*', '602', rev, packing_num, opn602_param, ',')
    rt = opn602_data.split(',')[1]
    pid = fits_dll.fn_query('*', '101', rev, rt, 'PID', ',')
    pass_kitting_qty = str(len(str(fits_dll.fn_query('*', '151', 'RT', '*', rt, ',')).split(',')))
    print('Pass Kitting Qty= {}'.format(pass_kitting_qty))
    result = 'Accept'
    print(opn602_data + ',' + pass_kitting_qty + ',' + pid + ',' + result)
    oba_data = opn602_data + ',' + pass_kitting_qty + ',' + pid + ',' + result
    fits_dll.closeDB
    return oba_data


def find_packing_num(fits_dll, rt):
    print("Input RT = {}".format(rt))
    if not fits_dll.fn_InitDB('*', rev, ''):
        return False
    # get & verify build type
    build_type = fits_dll.fn_query('*', '902', rev, rt,'Build Type')
    if build_type == 'DTS':
        return build_type
    sn_list = str(fits_dll.fn_query('*', '151', 'RT', '*', rt, ','))
    sn = sn_list.split(',')
    print('RT: {}, Quantity = {}'.format(rt, len(sn)))
    packing_num_list = []
    for i in range(len(sn)):
        # get packing numbers
        packing_num = fits_dll.fn_query('*', '601_B', rev, sn[i], "Packing No", ',')
        packing_num_list.append(packing_num)
    # find unique packing number
    list_of_unique_num = []
    unique_num = set(packing_num_list)
    for num in unique_num:
        if num != '-':
            list_of_unique_num.append(num)
    print('Unique packing number: {}'.format(list_of_unique_num))
    fits_dll.closeDB
    return list_of_unique_num


def save_opn702(fits_dll, rt, inv_num, packing_num, packing_qty):
    if not fits_dll.fn_InitDB('*', rev, ''):
        return False
    sn_list = str(get_sn_list(fits_dll, rt))
    for sn in sn_list.split(','):
        last_opn = get_last_opn(fits_dll, sn)
        print('SN: {} = Last operation: {}'.format(sn, last_opn))
        validate_route = valid_inv('702', sn)
        if last_opn == '601_B' and validate_route['status'] == True:
            data = '026487,' + inv_num + ',' + sn + ',' + packing_num + ',' + packing_qty
            print(data)
    fits_dll.closeDB
    return True


if __name__ == '__main__':
    RT = '2222221'
    packing_num = 'P210310015'
    packing_qty = '3'
    inv_num = '7777777'
    etr = 'R20210130'
    # Yanee
    yanee_en = "EN:507207"
    # Worawan
    worawan_en = "EN:507837"
    # Tarika
    tarika_en = "EN:509454"
    # Attapon
    attapon_en = "EN:026487"
    # Kreetha
    kreetha_en = "EN:015457"
    print(en_validate('1501', tarika_en))
    packing_number = find_packing_num(fits_dll ,RT)
    prepare_oba_info(fits_dll, packing_num)
    save_opn702(fits_dll ,RT, inv_num, packing_num, packing_qty)
    prepare_etr_info(fits_dll, etr)
    result = valid_inv('702', 'LCC2424A3DS')
    print(result['status'])
    print(type(result['status']))
    if result['status']:
        print('Allow')
    else:
        print('Not-Allow')
    print(get_last_opn(fits_dll ,inv_num))