#-*- encoding=UTF-8 -*-

__author__ = 'hedingbo@saicmotor.com'

import xml.etree.cElementTree as ET
import sys
import xmltodict
from collections import OrderedDict
import pandas as pd
import xlwings as xw

excel_header = ['Frame Name', 'CAN ID(Hex)', 'Frame Length(Byte)', 'Container Name', 'Container Length(Byte)', 'PDU Name', 'PDU ID(Hex)', 'PDU Length(Byte)', 'PDU Timing(ms)', 'Signal Short Name', 'Start Bit(MSB)', 'Length(Bit)', 'Signal Type', 'Init Value(Dec)', 'Conversion', 'CANFD Support', 'Direction']

# Define item loction in excel
_loc_frame_name = 1
_loc_can_id = 2
_loc_frame_length = 3
_loc_container_name = 4
_loc_container_length = 5
_loc_pdu_name = 6
_loc_pdu_id = 7
_loc_pdu_length = 8
_loc_pdu_timing = 9
_loc_signal_name = 10
_loc_start_bit = 11
_loc_signal_length = 12
_loc_signal_type = 13
_loc_init_value = 14
_loc_signal_conversion = 15
_loc_canfd_support = 16
_loc_frame_direction = 17

def dec2hex(string_num):
    if string_num == '0':
        return '0'
    else:
        base = [str(x) for x in range(10)] + [ chr(x) for x in range(ord('A'),ord('A')+6)]
        num = int(string_num)
        mid = []
        while True:
            if num == 0: break
            num,rem = divmod(num, 16)
            mid.append(base[rem])
        return ''.join([str(x) for x in mid[::-1]])

def create_dict_from_list(keys_list, items_list):
    dic = dict.fromkeys(keys_list, 0)
    for i in range(len(keys_list)):
        dic[keys_list[i]] = items_list[i]
    return dic

class ArxmlToXls:
    def __init__(self, arxmlFile):
        self.arxmlFile = arxmlFile
        self.topology = []
        self.frameList = []
        self.containerList = []
        self.pduList = []
        self.signalList = []
        self.arxml_to_dict()
    
    def arxml_to_dict(self):
        fd = open(self.arxmlFile)
        print(">>> Processing {}".format(self.arxmlFile))
        self.arxml_dict = xmltodict.parse(fd.read())
        self.parse_topology()
        self.parse_commatrix()
    
    def parse_topology(self):
        templist = []
        topolist = []
        if isinstance(self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][2]['ELEMENTS']['CAN-CLUSTER'], list):
            for cluster in self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][2]['ELEMENTS']['CAN-CLUSTER']:
                topodict = {}
                templist = []
                for frame in cluster['CAN-CLUSTER-VARIANTS']['CAN-CLUSTER-CONDITIONAL']['PHYSICAL-CHANNELS']['CAN-PHYSICAL-CHANNEL']['FRAME-TRIGGERINGS']['CAN-FRAME-TRIGGERING']:
                    tempdict = {}
                    tempdict['CANID'] = '0x'+dec2hex(frame['IDENTIFIER'])
                    tempdict['Cluster'] = cluster['SHORT-NAME']
                    tempdict['FrameName'] = frame['FRAME-REF']['#text'].split('/')[3]
                    tempdict['PDU'] = frame['PDU-TRIGGERINGS']['PDU-TRIGGERING-REF-CONDITIONAL']['PDU-TRIGGERING-REF']['#text'].split('/')[4][5:]
                    if frame['FRAME-PORT-REFS']['FRAME-PORT-REF']['#text'].split('/')[7] == 'FramePort_In':
                        tempdict['Direction'] = 'Recieve'
                        if frame['CAN-FRAME-RX-BEHAVIOR'] == 'CAN-FD':
                            tempdict['CANFD'] = 'True'
                        else:
                            tempdict['CANFD'] = 'False'
                    if frame['FRAME-PORT-REFS']['FRAME-PORT-REF']['#text'].split('/')[7] == 'FramePort_Out':
                        tempdict['Direction'] = 'Send'
                        if frame['CAN-FRAME-TX-BEHAVIOR'] == 'CAN-FD':
                            tempdict['CANFD'] = 'True'
                        else:
                            tempdict['CANFD'] = 'False'
                    
                    for pdutrig in cluster['CAN-CLUSTER-VARIANTS']['CAN-CLUSTER-CONDITIONAL']['PHYSICAL-CHANNELS']['CAN-PHYSICAL-CHANNEL']['PDU-TRIGGERINGS']['PDU-TRIGGERING']:
                        if tempdict['PDU'] == pdutrig['I-PDU-REF']['#text'].split('/')[3]:
                            if pdutrig['I-PDU-REF']['@DEST'] == 'CONTAINER-I-PDU':
                                tempdict['PDUType'] = 'Container'
                            elif pdutrig['I-PDU-REF']['@DEST'] == 'I-SIGNAL-I-PDU':
                                tempdict['PDUType'] = 'I-PDU'
                            else:
                                tempdict['PDUType'] = 'NM_TP'
                    templist.append(tempdict)
                topodict = {'Cluster': cluster['SHORT-NAME'],'Topology':templist}
                topolist.append(topodict)
        else:
            cluster = self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][2]['ELEMENTS']['CAN-CLUSTER']
            topodict = {}
            templist = []
            for frame in cluster['CAN-CLUSTER-VARIANTS']['CAN-CLUSTER-CONDITIONAL']['PHYSICAL-CHANNELS']['CAN-PHYSICAL-CHANNEL']['FRAME-TRIGGERINGS']['CAN-FRAME-TRIGGERING']:
                tempdict = {}
                tempdict['CANID'] = '0x'+dec2hex(frame['IDENTIFIER'])
                tempdict['Cluster'] = cluster['SHORT-NAME']
                tempdict['FrameName'] = frame['FRAME-REF']['#text'].split('/')[3]
                tempdict['PDU'] = frame['PDU-TRIGGERINGS']['PDU-TRIGGERING-REF-CONDITIONAL']['PDU-TRIGGERING-REF']['#text'].split('/')[4][5:]
                if frame['FRAME-PORT-REFS']['FRAME-PORT-REF']['#text'].split('/')[7] == 'FramePort_In':
                    tempdict['Direction'] = 'Recieve'
                    if frame['CAN-FRAME-RX-BEHAVIOR'] == 'CAN-FD':
                        tempdict['CANFD'] = 'True'
                    else:
                        tempdict['CANFD'] = 'False'
                if frame['FRAME-PORT-REFS']['FRAME-PORT-REF']['#text'].split('/')[7] == 'FramePort_Out':
                    tempdict['Direction'] = 'Send'
                    if frame['CAN-FRAME-TX-BEHAVIOR'] == 'CAN-FD':
                        tempdict['CANFD'] = 'True'
                    else:
                        tempdict['CANFD'] = 'False'
                
                for pdutrig in cluster['CAN-CLUSTER-VARIANTS']['CAN-CLUSTER-CONDITIONAL']['PHYSICAL-CHANNELS']['CAN-PHYSICAL-CHANNEL']['PDU-TRIGGERINGS']['PDU-TRIGGERING']:
                    if tempdict['PDU'] == pdutrig['I-PDU-REF']['#text'].split('/')[3]:
                        if pdutrig['I-PDU-REF']['@DEST'] == 'CONTAINER-I-PDU':
                            tempdict['PDUType'] = 'Container'
                        elif pdutrig['I-PDU-REF']['@DEST'] == 'I-SIGNAL-I-PDU':
                            tempdict['PDUType'] = 'I-PDU'
                        else:
                            tempdict['PDUType'] = 'NM_TP'
                templist.append(tempdict)
            topodict = {'Cluster': cluster['SHORT-NAME'],'Topology':templist}
            topolist.append(topodict)
        self.topology = topolist

    def parse_commatrix(self):
        tempframelist = []
        tempcontainerlist = []
        temppdulist = []
        tempsignallist = []
        temproutingtablelist = []

        # Initiate index list
        frame_index_dict = []
        cpdu_index_dict = []
        pdu_index_dict = []
        signal_index_dict = []

        # frame dict in arxml
        framepkg = self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][3]['AR-PACKAGES']['AR-PACKAGE'][0]

        # pdu dict in arxml
        pdupkg = self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][3]['AR-PACKAGES']['AR-PACKAGE'][1]

        # signal dict in arxml
        signalpkg = self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][3]['AR-PACKAGES']['AR-PACKAGE'][2]

        # Get signalgroup
        GwFlag = False
        for arpkg in self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][3]['AR-PACKAGES']['AR-PACKAGE']:
            if arpkg['SHORT-NAME'] == 'Gateway':
                routingpkg = arpkg
                GwFlag = True

        compupkg = self.arxml_dict['AUTOSAR']['AR-PACKAGES']['AR-PACKAGE'][5]['AR-PACKAGES']['AR-PACKAGE']['ELEMENTS']['COMPU-METHOD']

        signal_name_list = [sig['SHORT-NAME'][2:] for sig in signalpkg['ELEMENTS']['I-SIGNAL']]

        compu_name_list = [compu['SHORT-NAME'] for compu in compupkg]

        # Get frame info from frame dict
        keys = []
        items = []
        idx = 0
        for frame in framepkg['ELEMENTS']['CAN-FRAME']:
            tempdict = {}
            tempdict['FrameName'] = frame['SHORT-NAME']
            tempdict['DLC'] = frame['FRAME-LENGTH']
            tempframelist.append(tempdict)
            items.append(idx)
            keys.append(frame['SHORT-NAME'])
            idx += 1
        frame_index_dict = create_dict_from_list(keys, items)
        # Get container pdu info from container pdu dict
        keys = []
        items = []
        idx = 0
        if 'CONTAINER-I-PDU' in pdupkg['ELEMENTS']:
            for pdu in pdupkg['ELEMENTS']['CONTAINER-I-PDU']:
                tempdict = {}
                contedlist = []
                tempdict['ContainerName'] = pdu['SHORT-NAME']
                tempdict['ContainerLength'] = pdu['LENGTH']
                if isinstance(pdu['CONTAINED-PDU-TRIGGERING-REFS']['CONTAINED-PDU-TRIGGERING-REF'], list):
                    for contPdu in pdu['CONTAINED-PDU-TRIGGERING-REFS']['CONTAINED-PDU-TRIGGERING-REF']:
                        contedlist.append(contPdu['#text'].split('/')[4][5:])
                else:
                    contedlist.append(pdu['CONTAINED-PDU-TRIGGERING-REFS']['CONTAINED-PDU-TRIGGERING-REF']['#text'].split('/')[4][5:])
                tempdict['ContainedPdus'] = contedlist
                items.append(idx)
                keys.append(pdu['SHORT-NAME'])
                tempcontainerlist.append(tempdict)
                idx += 1
            cpdu_index_dict = create_dict_from_list(keys, items)

        # Get pdu info from pdu dict
        pdukeys = []
        pduitems = []
        pduidx = 0
        signalkeys = []
        signalitems = []
        signalidx = 0
        for pdu in pdupkg['ELEMENTS']['I-SIGNAL-I-PDU']:
            tempdict = {}
            tempdict['IPduName'] = pdu['SHORT-NAME']
            if 'CONTAINED-I-PDU-PROPS' in pdu.keys():
                tempdict['PduID'] = '0x'+dec2hex(pdu['CONTAINED-I-PDU-PROPS']['HEADER-ID-SHORT-HEADER'])
            else:
                tempdict['PduID'] = 'None'
            tempdict['PduLength'] = pdu['LENGTH']
            # Get pdu timing parameter
            if 'I-PDU-TIMING-SPECIFICATIONS' in pdu.keys():
                if 'EVENT-CONTROLLED-TIMING' in pdu['I-PDU-TIMING-SPECIFICATIONS']['I-PDU-TIMING']['TRANSMISSION-MODE-DECLARATION']['TRANSMISSION-MODE-TRUE-TIMING'].keys():
                    tempdict['Timing'] = 'Sporadic'
                else:
                    tempdict['Timing'] = str(float(pdu['I-PDU-TIMING-SPECIFICATIONS']['I-PDU-TIMING']['TRANSMISSION-MODE-DECLARATION']['TRANSMISSION-MODE-TRUE-TIMING']['CYCLIC-TIMING']['TIME-PERIOD']['VALUE'])*1000) 
            else:
                tempdict['Timing'] = 'None'
            # pdu indexing
            pdukeys.append(pdu['SHORT-NAME'])
            pduitems.append(pduidx)
            pduidx += 1

            # Get signals info in current pdu
            mappedsignals = []
            if 'I-SIGNAL-TO-PDU-MAPPINGS' in pdu.keys():
                if isinstance(pdu['I-SIGNAL-TO-PDU-MAPPINGS']['I-SIGNAL-TO-I-PDU-MAPPING'], list):
                    for signal in pdu['I-SIGNAL-TO-PDU-MAPPINGS']['I-SIGNAL-TO-I-PDU-MAPPING']:
                        if 'I-SIGNAL-GROUP-REF' in signal.keys():
                            continue 
                        else:
                            signaldict = {}
                            signaldict['MappedPdu'] = pdu['SHORT-NAME']
                            signaldict['SignalName'] = signal['I-SIGNAL-REF']['#text'].split('/')[3][2:]
                            signaldict['StartBit'] = signal['START-POSITION']
                            signaldict['DataType'] = 'None'
                            signaldict['SignalLength'] = '0'
                            signaldict['InitValue'] = '0'
                            signaldict['Conversion'] = 'None'
                            tempsignallist.append(signaldict)
                            mappedsignals.append(signaldict['SignalName'])
                            signalkeys.append(signaldict['SignalName'])
                            signalitems.append(signalidx)
                            signalidx += 1
                else:
                    signal = pdu['I-SIGNAL-TO-PDU-MAPPINGS']['I-SIGNAL-TO-I-PDU-MAPPING']
                    signaldict = {}
                    signaldict['MappedPdu'] = pdu['SHORT-NAME']
                    signaldict['SignalName'] = signal['I-SIGNAL-REF']['#text'].split('/')[3][2:]
                    signaldict['StartBit'] = signal['START-POSITION']
                    signaldict['DataType'] = 'None'
                    signaldict['SignalLength'] = '0'
                    signaldict['InitValue'] = '0'
                    signaldict['Conversion'] = 'None'
                    mappedsignals.append(signaldict['SignalName'])
                    tempsignallist.append(signaldict)
                    signalkeys.append(signaldict['SignalName'])
                    signalitems.append(signalidx)
                    signalidx += 1

                tempdict['MapSignals'] = mappedsignals
                temppdulist.append(tempdict)
            else:
                tempdict['MapSignals'] = ['PDUGateway']
                temppdulist.append(tempdict)

        pdu_index_dict = create_dict_from_list(pdukeys, pduitems)
        signal_index_dict = create_dict_from_list(signalkeys, signalitems)
        # Get signal info from signal dict
        for elem in tempsignallist:
            isignal = signalpkg['ELEMENTS']['I-SIGNAL'][signal_name_list.index(elem['SignalName'])]
            elem['DataType'] = isignal['NETWORK-REPRESENTATION-PROPS']['SW-DATA-DEF-PROPS-VARIANTS']['SW-DATA-DEF-PROPS-CONDITIONAL']['BASE-TYPE-REF']['#text'].split('/')[4]
            elem['SignalLength'] = isignal['LENGTH']
            if 'NUMERICAL-VALUE-SPECIFICATION' in isignal['INIT-VALUE'].keys():
                elem['InitValue'] = isignal['INIT-VALUE']['NUMERICAL-VALUE-SPECIFICATION']['VALUE']
            elif 'ARRAY-VALUE-SPECIFICATION' in  isignal['INIT-VALUE'].keys():
                init_value = ''
                for value in isignal['INIT-VALUE']['ARRAY-VALUE-SPECIFICATION']['ELEMENTS']['NUMERICAL-VALUE-SPECIFICATION']:
                    init_value = init_value + value['VALUE']+','
                elem['InitValue'] = '['+init_value[:-1]+']'
            elem['Conversion'] = self.parse_compu_method(compupkg[compu_name_list.index(isignal['NETWORK-REPRESENTATION-PROPS']['SW-DATA-DEF-PROPS-VARIANTS']['SW-DATA-DEF-PROPS-CONDITIONAL']['COMPU-METHOD-REF']['#text'].split('/')[3])])
        
        # Get routing path table for gateway node(Only valid for gateway-enabled node)
        if GwFlag:
            pduGwPath = routingpkg['ELEMENTS']['GATEWAY']['I-PDU-MAPPINGS']['I-PDU-MAPPING']
            sigGwPath = routingpkg['ELEMENTS']['GATEWAY']['SIGNAL-MAPPINGS']['I-SIGNAL-MAPPING']
            for pdugw in pduGwPath:
                pathdict = {}
                pathdict['Name'] = pdugw['SOURCE-I-PDU-REF']['#text'].split('/')[4][5:]
                pathdict['Type'] = 'PDU Gateway'
                pathdict['Source'] = pdugw['SOURCE-I-PDU-REF']['#text'].split('/')[2]
                pathdict['Target'] = pdugw['TARGET-I-PDU']['TARGET-I-PDU-REF']['#text'].split('/')[2]
                temproutingtablelist.append(pathdict)
            
            for siggw in sigGwPath:
                pathdict = {}
                pathdict['Name'] = siggw['SOURCE-SIGNAL-REF']['#text'].split('/')[4][6:-4]
                pathdict['Type'] = 'Signal Gateway'
                pathdict['Source'] = siggw['SOURCE-SIGNAL-REF']['#text'].split('/')[2]
                pathdict['Target'] = siggw['TARGET-SIGNAL-REF']['#text'].split('/')[2]
                temproutingtablelist.append(pathdict)

        self.frameList = tempframelist
        self.containerList = tempcontainerlist
        self.pduList = temppdulist
        self.signalList = tempsignallist
        self.routingList = temproutingtablelist
        self.gwflag = GwFlag

        self.frameIndex = frame_index_dict
        self.cpduIndex = cpdu_index_dict
        self.pduIndex = pdu_index_dict
        self.signalIndex = signal_index_dict

    def parse_compu_method(self, compuobj):
        conversion = ''
        # Get compu-mehtod dict
        if compuobj['CATEGORY'] == 'TEXTTABLE':
            if 'COMPU-INTERNAL-TO-PHYS' in compuobj.keys():
                if isinstance(compuobj['COMPU-INTERNAL-TO-PHYS']['COMPU-SCALES']['COMPU-SCALE'], list):
                    for compu_scale in compuobj['COMPU-INTERNAL-TO-PHYS']['COMPU-SCALES']['COMPU-SCALE']:
                        tempstr = ''
                        tempstr = '0x'+dec2hex(compu_scale['LOWER-LIMIT']['#text'])+'='+ compu_scale['COMPU-CONST']['VT']
                        conversion = conversion + tempstr +'\n'
                    conversion = conversion[:-1]
                else:
                    conversion = '0x'+dec2hex(compuobj['COMPU-INTERNAL-TO-PHYS']['COMPU-SCALES']['COMPU-SCALE']['LOWER-LIMIT']['#text'])+'='+ compuobj['COMPU-INTERNAL-TO-PHYS']['COMPU-SCALES']['COMPU-SCALE']['COMPU-CONST']['VT']
            else:
                # for compu_scale in compuobj['COMPU-PHYS-TO-INTERNAL']['COMPU-SCALES']['COMPU-SCALE']:
                #     tempstr = ''
                #     tempstr = dec2hex(compu_scale['LOWER-LIMIT']['#text'])+'='+ compu_scale['COMPU-CONST']['VT']
                #     conversion = conversion + tempstr +'\n'
                conversion = 'INF = Unknown'

        elif compuobj['CATEGORY'] == 'LINEAR':
            compu_scale = compuobj['COMPU-INTERNAL-TO-PHYS']['COMPU-SCALES']['COMPU-SCALE']
            if float(compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]) == 0:
                conversion = 'E='+'N*'+compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][1]
            elif float(compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]) > 0:
                conversion = 'E='+'N*'+compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][1] + '+' + compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]
            else:
                conversion = 'E='+'N*'+compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][1] + compu_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]

        elif compuobj['CATEGORY'] == 'SCALE_LINEAR_AND_TEXTTABLE':
            linear_scales = []
            txt_scales = []
            for compu_scale in compuobj['COMPU-INTERNAL-TO-PHYS']['COMPU-SCALES']['COMPU-SCALE']:
                if 'COMPU-RATIONAL-COEFFS' in compu_scale.keys():
                    linear_scales.append(compu_scale)
                else:
                    txt_scales.append(compu_scale)

            for linear_scale in linear_scales:
                tempstr = ''
                if float(linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]) == 0:
                    tempstr = 'E='+'N*'+linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][1]
                elif float(linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]) > 0:
                    tempstr = 'E='+'N*'+linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][1] + '+' + linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]
                else:
                    tempstr = 'E='+'N*'+linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][1] + linear_scale['COMPU-RATIONAL-COEFFS']['COMPU-NUMERATOR']['V'][0]
                conversion = conversion + tempstr + '\n'

            for txt_scale in txt_scales:
                tempstr = ''
                tempstr = '0x'+dec2hex(txt_scale['LOWER-LIMIT']['#text'])+'='+ txt_scale['COMPU-CONST']['VT']
                conversion = conversion + tempstr +'\n'
            conversion = conversion[:-1]
        else:
            conversion = 'None'

        return conversion

    def write_arxml_to_excel(self):
        # Create new workbook
        wb = xw.Book()

        for cluster in self.topology: 
            sht = wb.sheets.add(cluster['Cluster'])
            
            # Write table header
            sht.range((1,1),(1,len(excel_header))).value = excel_header

            # Set header format
            header_rng = sht.range((1,1),(1,len(excel_header)))
            header_rng.color = 217,217,217
            # header_rng.api.Font.Name = 'Arial'
            header_rng.api.Font.ColorIndex = 1
            header_rng.api.Font.Size = 9
            header_rng.api.Font.Bold = True
            header_rng.api.HorizontalAlignment = -4108
            header_rng.api.VerticalAlignment = -4107
            header_rng.api.Orientation = -4171
            # Internal
            header_rng.api.Borders(11).LineStyle = 1
            header_rng.api.Borders(11).Weight = 2
            # Left
            header_rng.api.Borders(7).LineStyle = 1
            header_rng.api.Borders(7).Weight = -4138
            # Top
            header_rng.api.Borders(8).LineStyle = 1
            header_rng.api.Borders(8).Weight = -4138
            # Right
            header_rng.api.Borders(10).LineStyle = 1
            header_rng.api.Borders(10).Weight = -4138

            row_counter = 2
            msg_counter = 2
            for frame in cluster['Topology']:
                sht.range(row_counter, _loc_can_id).value = frame['CANID']
                sht.range(row_counter, _loc_frame_name).value = frame['FrameName']
                sht.range(row_counter, _loc_frame_length).value = int(self.frameList[self.frameIndex[frame['FrameName']]]['DLC'])
                sht.range(row_counter, _loc_frame_direction).value = frame['Direction']
                sht.range(row_counter, _loc_canfd_support).value = frame['CANFD']

                row_counter += 1
                if frame['PDUType'] == 'Container':
                    sht.range(row_counter, _loc_container_name).value = frame['PDU']
                    cpdu = self.containerList[self.cpduIndex[frame['PDU']]]
                    sht.range(row_counter, _loc_container_length).value = int(cpdu['ContainerLength'])
                    # Set frame format
                    sht.range((msg_counter,_loc_frame_name), (row_counter, _loc_container_length)).color = 0,255,255

                    row_counter += 1
                    for elem in cpdu['ContainedPdus']:
                        pdu = self.pduList[self.pduIndex[elem]]
                        sht.range(row_counter, _loc_pdu_name).value = pdu['IPduName']
                        sht.range(row_counter, _loc_pdu_id).value = pdu['PduID']
                        sht.range(row_counter, _loc_pdu_timing).value = pdu['Timing']
                        sht.range(row_counter, _loc_pdu_length).value = pdu['PduLength']
                        sht.range((row_counter, _loc_pdu_name), (row_counter, _loc_pdu_timing)).color = 255,255,0
                        row_counter += 1
                        # Write signals to excel
                        for item in pdu['MapSignals']:
                            if item == 'PDUGateway':
                                sht.range(row_counter, _loc_signal_name).value = 'PDU Gateway'
                                sht.range(row_counter, _loc_signal_name).color = 102,255,102
                                sht.range(row_counter, _loc_signal_name).api.Font.Italic = True                                
                                row_counter += 1
                            else:
                                signal = self.signalList[self.signalIndex[item]]
                                sht.range(row_counter, _loc_signal_name).value = signal['SignalName']
                                sht.range(row_counter, _loc_start_bit).value = signal['StartBit']
                                sht.range(row_counter, _loc_signal_length).value = signal['SignalLength']
                                sht.range(row_counter, _loc_signal_type).value = signal['DataType']
                                sht.range(row_counter, _loc_init_value).value = signal['InitValue']
                                sht.range(row_counter, _loc_signal_conversion).value = signal['Conversion']
                                sht.range((row_counter, _loc_signal_name), (row_counter, _loc_frame_direction)).color = 255,255,0
                                row_counter += 1
                        
                elif frame['PDUType'] == 'I-PDU':
                    sht.range(row_counter, _loc_container_name).value = 'None'
                    sht.range(row_counter, _loc_container_length).value = 0
                    sht.range((msg_counter,_loc_frame_name), (row_counter, _loc_container_length)).color = 0,255,255
                    # Write pdu info
                    row_counter += 1
                    pdu = self.pduList[self.pduIndex[frame['PDU']]]
                    sht.range(row_counter, _loc_pdu_name).value = pdu['IPduName']
                    sht.range(row_counter, _loc_pdu_id).value = pdu['PduID']
                    sht.range(row_counter, _loc_pdu_timing).value = pdu['Timing']
                    sht.range(row_counter, _loc_pdu_length).value = pdu['PduLength']
                    sht.range((row_counter, _loc_pdu_name), (row_counter, _loc_pdu_timing)).color = 255,255,0
                    row_counter += 1

                    # Write signals to excel
                    for item in pdu['MapSignals']:
                        if item == 'PDUGateway':
                            sht.range(row_counter, _loc_signal_name).value = 'PDU Gateway'
                            sht.range(row_counter, _loc_signal_name).color = 102,255,102
                            sht.range(row_counter, _loc_signal_name).api.Font.Italic = True
                            row_counter += 1
                        else:
                            signal = self.signalList[self.signalIndex[item]]
                            sht.range(row_counter, _loc_signal_name).value = signal['SignalName']
                            sht.range(row_counter, _loc_start_bit).value = signal['StartBit']
                            sht.range(row_counter, _loc_signal_length).value = signal['SignalLength']
                            sht.range(row_counter, _loc_signal_type).value = signal['DataType']
                            sht.range(row_counter, _loc_init_value).value = signal['InitValue']
                            sht.range(row_counter, _loc_signal_conversion).value = signal['Conversion']
                            sht.range((row_counter, _loc_signal_name), (row_counter, _loc_frame_direction)).color = 255,255,0
                            row_counter += 1
                else:
                    sht.range(row_counter, _loc_container_name).value = 'None'
                    sht.range(row_counter, _loc_container_length).value = 0
                    sht.range((msg_counter,_loc_frame_name), (row_counter, _loc_container_length)).color = 0,255,255
                    # Write pdu info
                    row_counter += 1
                    sht.range(row_counter, _loc_pdu_name).value = frame['PDU']
                    sht.range(row_counter, _loc_pdu_id).value = 'None'
                    sht.range(row_counter, _loc_pdu_timing).value = 'None'
                    sht.range(row_counter, _loc_pdu_length).value = 'None'
                    sht.range((row_counter, _loc_pdu_name), (row_counter, _loc_pdu_timing)).color = 255,255,0
                    row_counter += 1
                
                # Set frame format
                frame_rng = sht.range((msg_counter, _loc_frame_name), (row_counter-1, _loc_frame_direction))
                # Internal borders: horizontal and vertical
                frame_rng.api.Borders(11).LineStyle = 1
                frame_rng.api.Borders(11).Weight = 2
                frame_rng.api.Borders(12).LineStyle = 1
                frame_rng.api.Borders(12).Weight = 2
                # top 
                frame_rng.api.Borders(8).LineStyle = 1
                frame_rng.api.Borders(8).Weight = -4138
                # bottom
                frame_rng.api.Borders(9).LineStyle = 1
                frame_rng.api.Borders(9).Weight = -4138
                # left
                frame_rng.api.Borders(7).LineStyle = 1
                frame_rng.api.Borders(7).Weight = -4138
                # right
                frame_rng.api.Borders(10).LineStyle = 1
                frame_rng.api.Borders(10).Weight = -4138
                msg_counter = row_counter

            # Set sheet format
            sht.autofit()
            sht.api.UsedRange.Font.Name = 'Arial'
            sht.api.UsedRange.Font.Size = 9
            sht.api.Range("B:B").ColumnWidth = 5
            sht.api.Range("C:C").ColumnWidth = 3
            sht.api.Range("E:E").ColumnWidth = 3
            sht.api.Range("G:G").ColumnWidth = 8
            sht.api.Range("H:H").ColumnWidth = 4
            sht.api.Range("I:I").ColumnWidth = 8
            sht.api.Range("K:K").ColumnWidth = 3
            sht.api.Range("L:L").ColumnWidth = 3
            sht.api.Range("M:M").ColumnWidth = 8.5
            sht.api.Range("N:N").ColumnWidth = 6
            sht.api.Range("P:P").ColumnWidth = 6
            sht.api.Range("Q:Q").ColumnWidth = 6

            # Set alignment
            sht.api.Range("B:B").HorizontalAlignment = -4108
            sht.api.Range("C:C").HorizontalAlignment = -4108
            sht.api.Range("E:E").HorizontalAlignment = -4108
            sht.api.Range("G:G").HorizontalAlignment = -4108
            sht.api.Range("H:H").HorizontalAlignment = -4108
            sht.api.Range("I:I").HorizontalAlignment = -4108
            sht.api.Range("K:K").HorizontalAlignment = -4108
            sht.api.Range("L:L").HorizontalAlignment = -4108
            sht.api.Range("M:M").HorizontalAlignment = -4108
            sht.api.Range("N:N").HorizontalAlignment = -4108
            sht.api.Range("P:P").HorizontalAlignment = -4108
            sht.api.Range("Q:Q").HorizontalAlignment = -4108
        
        # Write routing table
        if self.gwflag:
            sht = wb.sheets['Sheet1']
            sht.api.Name = 'Routing Table'
            sht.api.Tab.ColorIndex = 4
            sht.range((1,1), (1,4)).value = ['Name', 'Routing Type', 'Source', 'Target']
            header_rng = sht.range((1,1),(1,4))
            header_rng.color = 217,217,217
            header_rng.api.Font.ColorIndex = 1
            # header_rng.api.Font.Size = 11
            header_rng.api.Font.Bold = True

            row_counter = 2
            for path in self.routingList:
                sht.range(row_counter, 1).value = path['Name']
                sht.range(row_counter, 2).value = path['Type']
                sht.range(row_counter, 3).value = path['Source']
                sht.range(row_counter, 4).value = path['Target']
                row_counter += 1

            # Format sheet
            sht.autofit()
            sht.api.UsedRange.Font.Name = 'Arial'
            sht.api.UsedRange.Font.Size = 9
            sht.api.UsedRange.HorizontalAlignment = -4108
            sht.api.UsedRange.VerticalAlignment = -4108
            sht.api.UsedRange.RowHeight = 20
            # Set border style
            sht.api.UsedRange.Borders(11).LineStyle = 1
            sht.api.UsedRange.Borders(11).Weight = 2
            sht.api.UsedRange.Borders(12).LineStyle = 1
            sht.api.UsedRange.Borders(12).Weight = 2
            # top 
            sht.api.UsedRange.Borders(8).LineStyle = 1
            sht.api.UsedRange.Borders(8).Weight = -4138
            # bottom
            sht.api.UsedRange.Borders(9).LineStyle = 1
            sht.api.UsedRange.Borders(9).Weight = -4138
            # left
            sht.api.UsedRange.Borders(7).LineStyle = 1
            sht.api.UsedRange.Borders(7).Weight = -4138
            # right
            sht.api.UsedRange.Borders(10).LineStyle = 1
            sht.api.UsedRange.Borders(10).Weight = -4138

        self.optFileName = self.arxmlFile[:-5]+'xlsx'
        wb.save(self.optFileName)
        wb.close()
        print(">>> Done!")

    def output_result(self): 
        '''
        ' For debug only
        '''
        optList = []
        
        df1 = pd.DataFrame(self.frameList)
        df2 = pd.DataFrame(self.containerList)
        df3 = pd.DataFrame(self.pduList)
        df4 = pd.DataFrame(self.signalList)
        df5 = pd.DataFrame(self.topology)

        df1.to_excel('framelist.xlsx', 'Frame', encoding='utf-8', index=False)
        df2.to_excel('containerlist.xlsx', 'Container', encoding='utf-8', index=False)
        df3.to_excel('pdulist.xlsx', 'Pdu', encoding='utf-8', index=False)
        df4.to_excel('signallist.xlsx', 'Signal', encoding='utf-8', index=False)
        df5.to_excel('topology.xlsx', 'topology', encoding='utf-8', index=False)

if __name__ == '__main__':
    # Load arxml file
    # arFile = r'History Files\20200728-hd2-EP33L_Simu1_RWS-V1.1-updatetype.arxml'
    # arxmlObject = ArxmlToXls(arFile)
    # arxmlObject.write_arxml_to_excel()
    pass



 