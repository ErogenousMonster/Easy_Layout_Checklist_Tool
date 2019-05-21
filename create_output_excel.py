# coding=utf-8

import xlwings as xw
import math
import copy
import os
import re
from tkinter import _flatten
import xlsxwriter
import sys


class brd_data(object):
    def __init__(self, data):
        self.data = data

    def GetData(self):
        return self.data


# 定义错误输出
def create_error_message(excel_path, error_message):
    # 創建excel
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet('error_message')

    error_format = workbook.add_format({'font_size': 22})

    worksheet.write('A1', 'Program running error:', error_format)
    worksheet.write('B2', error_message + ', please check and try again!', error_format)

    workbook.close()


def get_parameter():
    """从parameter.txt中获取参数"""

    input_path = os.getcwd()
    param_path = os.path.join(input_path, 'parameter.txt')
    output_path = os.path.join('\\'.join(input_path.split('\\')[:-1]), 'output')
    excel_path = os.path.join(output_path, 'checklist.xlsx')
    try:
        parameter_file = open(param_path, 'r', encoding='latin-1', errors='ignore')
        parameter = parameter_file.read().strip('\n').split('\x9a')[-1]
        parameter_file.close()
        # parameter_list = parameter_list.split('\n')
        return parameter, param_path, output_path, excel_path
    except FileNotFoundError:
        error_message = 'No parameters filled in'
        create_error_message(excel_path, error_message)
        raise FileNotFoundError
    # except UnicodeDecodeError:
    #     error_message = 'No parameters filled in'
    #     create_error_message(excel_path, error_message)
    #     raise UnicodeDecodeError


def LoadAllegroFile():
    """通过获取的参数生成报告内容"""
    # 报告指令
    report_list = ['eld_rep.rpt', 'dpg_rep.rpt', 'spn_rep.rpt', 'xSection_rep.rpt']
    # command_list = ['eld', 'dpg', 'spn', 'x-section']

    # 报告路径(四个报告)
    input_path = os.getcwd()
    # print('input_path', input_path)
    flow_number_path = '\\'.join(input_path.split('\\')[-3:])
    user_path = os.path.join('\\'.join(os.getcwd().split('\\')[:-4]), 'UserUpload')
    rpt_path = os.path.join(user_path, flow_number_path)
    excel_path = os.path.join('\\'.join(os.getcwd().split('\\')[:-1]), 'output\\checklist.xlsx')
    # print('rpt_path', rpt_path)

    report_path_list = [os.path.join(rpt_path, x) for x in report_list]

    # 读取生成报告中的数据
    try:
        data_list = [open(path, 'r').read() for path in report_path_list]
    except FileNotFoundError as e:
        error_message = 'Missing {} file in upload file'.format(str(e).split('\\')[-1][:-1])
        create_error_message(excel_path, error_message)
        raise FileNotFoundError

    # 数据整理，将有用信息提取出，两种不同的提取方法，适用于不同格式的报告
    content = ''
    for d in data_list:
        # 对走线报告进行分析
        if 'Detailed Etch Length by Layer and Width Report' in d or \
                'Detailed Trace Length by Layer and Width Report' in d:
            d = d.split('\n')
            d_tmp = list(map(lambda x: x.split(','), d[5:-1]))
            xy_point_pattern = re.compile(r'.*?xprobe:xy:\((.*?)\).*?xprobe:xy:\((.*?)\).*$')
            for idx in range(len(d_tmp)):
                xy_point = re.findall(xy_point_pattern, d_tmp[idx][-1])
                d_tmp[idx][-1] = ' '.join([xy_point[0][0], xy_point[0][1]])
            d_tmp = list(map(lambda x: ','.join(x), d_tmp))
            d = '\n'.join(d[0:5] + d_tmp) + '\n'
        # 对快速报告进行分析
        elif 'Allegro Report' in d:  # or 'Diffpair Gap Report' in d:
            d = d.split('\n')
            d_tmp = d[5:-1]
            d_tmp_ = list()
            for str1 in d_tmp:
                # find方法寻找字符串，找到返回位置索引，找不到返回-1
                if str1.find('"') > -1:
                    d_tmp_.append(','.join(re.split('\s\(|\)",|,\s', str1)[1:3]))
                else:
                    d_tmp_.append(','.join(re.split(',', str1)[0:1]))
            # 分离差分对名称
            d_tmp_ = sorted(list(set(d_tmp_)))
            d = '\n'.join(d[0:5] + d_tmp_) + '\n'
        content += d

    report_content = content

    return report_content


def report_filtrate():

    content = LoadAllegroFile()
    content = re.split('Detailed Etch Length by Layer and Width Report\n|Detailed Trace Length by Layer and Width '
                       'Report\n|Symbol Pin Report\n|Allegro Report\n|Cross Section Report\n', content)
    content = [x.split('\n') for x in content if x != '']

    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data = None, None, None, None
    for item in content:
        # print(item[3])
        if item[3].find('REFDES,PIN_NUMBER,SYM_NAME,COMP_DEVICE_TYPE,PAD_STACK_NAME,PIN_X,PIN_Y,NET_NAME') > -1:
            SCH_brd_data = brd_data([x.upper().split(',') for x in item[4:-1]])
        elif item[3].find(
                'Net Name,Total Net Length (mils),Layer Name,Total Layer Length (mils),Layer Length % '
                'of Total,Line Width (mils),Contiguous Length at Width (mils),Contiguous Length % Layer Length,'
                'Contiguous Length End Points') > -1:
            DEL_content = [x.upper().split(',') for x in item[4:-1]]
            for ind1 in range(len(DEL_content)):
                xy_point = DEL_content[ind1][-1].split(' ')
                # print(xy_point)
                DEL_content[ind1][-1] = [tuple(xy_point[0:2]), tuple(xy_point[2::])]
            Net_brd_data = brd_data(DEL_content)
        elif item[3].find(
                'Diffpair (Nets),Nominal Gap (mils),Actual Gap (mils),Gap Deviation (mils),Segment Length (mils),'
                'Segment End Points') > -1:
            diff_pair_brd_data = brd_data([x.upper().split(',') for x in item[4:-1]])
        elif item[4].find(
                'Subclass Name,Type,Material,Thickness (MIL),Conductivity (mho/cm),Dielectric Constant,Loss Tangent,'
                'Negative Artwork,Shield,Width (MIL),Unused Pin Pad Suppression,Unused Via Pad Suppression') > -1:
            Stackup_content = [x.upper() for x in item[5:-1]]
            stackup_brd_data = brd_data(Stackup_content)
        elif item[4].find(
                'Subclass Name,Type,Material,Thickness (MIL),Tol +,Tol -,Conductivity (mho/cm),Dielectric Constant,'
                'Loss Tangent,Negative Artwork,Shield,Width (MIL),Unused Pin Pad Suppression,Unused Via Pad Suppression'
        ) > -1:
            Stackup_content = [','.join(x[0:4] + x[6::]) for x in map(lambda y: y.upper().split(','), item[5:-1])]
            stackup_brd_data = brd_data(Stackup_content)
        elif item[4].find(
                'Subclass Name,Type,Material,Thickness (MIL),Conductivity (mho/cm),Dielectric Constant,Loss Tangent,'
                'Negative Artwork,Shield,Width (MIL),Single Impedance (ohm),Unused Pin Pad Suppression,'
                'Unused Via Pad Suppression') > -1:
            Stackup_content = [','.join(x[0:9] + x[10::]) for x in map(lambda y: y.upper().split(','), item[5:-1])]
            stackup_brd_data = brd_data(Stackup_content)

    if SCH_brd_data and Net_brd_data and diff_pair_brd_data and stackup_brd_data:
        return SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data


def total_mismatch(start_net_list, net_total_length_list):
    # print(len(start_net_list))
    # 仅针对差分信号线
    net_total_mismatch_list = []
    if len(start_net_list) >= 2 and len(start_net_list) % 2 == 0:
        for idx in range(int(len(start_net_list)/2)):
            # start_sch1, start_net1, end_sch1 = start_sch_list[2*idx], start_net_list[2*idx], end_sch_list[2*idx]
            # start_sch2, start_net2, end_sch2 = start_sch_list[2*idx+1], start_net_list[2*idx+1], end_sch_list[2*idx+1]

            total_length1 = net_total_length_list[2*idx]
            total_length2 = net_total_length_list[2*idx + 1]

            len1 = float(total_length1)
            len2 = float(total_length2)

            net_total_mismatch_list.append(round(abs(len1 - len2), 2))
            net_total_mismatch_list.append(round(abs(len1 - len2), 2))
    else:
        net_total_mismatch_list.append('NA')

    return net_total_mismatch_list


def get_total_signal(net_name_list, topology_dict):
    net_total_length_list = []
    net_total_via_list = []
    net_total_signal_list = []
    net_judge_total_signal = []
    # 读取所选信号线的Total length
    # show出所有经过的信号线
    for net in net_name_list:
        net_total_half_signal = ''
        net_judge_half_signal = []
        if topology_dict.get(net):
            net_half_list = topology_dict.get(net)
            net_half_list.sort(key=lambda x: len(x))
            net_total_length_list.append(round(float(net_half_list[-1][4].split(' ')[-1]), 2))
            net_total_via_list.append(int(net_half_list[-1][3].split(' ')[-1]))
            net_total_half_signal = net
            net_judge_half_signal.append(net)
            for x in net_half_list[-1]:
                if str(x).find('net$') > -1:
                    # net_total_half_signal.append(str(x).split('$')[-1])
                    net_total_half_signal += ' --> ' + str(x).split('$')[-1]
                    net_judge_half_signal.append(str(x).split('$')[-1])

        net_total_signal_list.append(net_total_half_signal)
        net_judge_total_signal.append(tuple(net_judge_half_signal))

    return net_total_signal_list, net_total_length_list, net_total_via_list, net_judge_total_signal


def get_one_way_signal(judge_signal_list, name_length_mismatch_dict, diff_pair_dict=None):
    # print(111, judge_signal_list)
    # print(222, name_length_mismatch_list)

    if diff_pair_dict:
        diff_pair_dict1 = {v: k for k, v in diff_pair_dict.items()}
        diff_pair_dict.update(diff_pair_dict1)

    name_length_mismatch_list = []
    for x in range(len(judge_signal_list)):
        item = list(copy.copy(judge_signal_list[x]))
        item.reverse()
        if tuple(item) in judge_signal_list[: x + 1] and len(item) > 1:
            # print(333, diff_pair_dict.get(judge_signal_list[x][0]))
            # print(444, judge_signal_list[x][-1])
            if diff_pair_dict and diff_pair_dict.get(judge_signal_list[x][0]) == judge_signal_list[x][-1]:
                # print('ok')
                name_length_mismatch_list.append(name_length_mismatch_dict.get(judge_signal_list[x]))
            pass
        else:
            name_length_mismatch_list.append(name_length_mismatch_dict.get(judge_signal_list[x]))
            # print(x)
            # print(item)
            # print(judge_signal_list[: x + 1])
            # print(name_length_mismatch_list[x])
            # del name_length_mismatch_list[x]

    # print(3333, name_length_mismatch_list)
    return name_length_mismatch_list


"""通过对分报告数据的处理，返回差分线，单根线，非信号线组"""


# 获得差分信号键值对，字典形式保存
def diff_detect(diff_pair_brd_data):
    # Get diff. pair list
    rpt_content = diff_pair_brd_data.GetData()
    both_diff = []
    lone_diff = []
    # print(rpt_content)
    for i in range(len(rpt_content)):
        if len(rpt_content[i]) > 1:
            both_diff.append(rpt_content[i])
        else:
            lone_diff.append(rpt_content[i])
    # 单向键值对
    diff_pair_dict = dict(both_diff)
    # 双向键值对
    # diff_pair_dict.update(dict([(x[1], x[0]) for x in both_diff]))
    # print(diff_pair_dict)

    # print(1111, diff_pair_dict)
    # print(2222, lone_diff)
    return diff_pair_dict, lone_diff


# 找出一个字符串中匹配某个字符的最后一个字符的索引
def find_last(string, str1):
    last_position = -1
    while True:
        position = string.find(str1, last_position + 1)
        if position == -1:
            return last_position
        last_position = position


# 获得电源线和地线的名称列表
def get_exclude_netlist(netlist):    # netlist = All_Net_List
    # Get pwr and gnd net list

    PWR_Net_KeyWord_List = ['^\+.*', '^-.*',
                            'VREF|VPP|VSS|PWR|VREG|VCORE|VCC|VT|VDD|VLED|PWM|VDIMM|VGT|VIN|[^S](VID)|VR',
                            'VOUT|VGG|VGPS|VNN|VOL|VSD|VSYS|VCM|VSA', '.*V[0-9]A.*', '.*V[0-9]\.[0-9]A.*',
                            '.*V[0-9]_[0-9]A.*', '.*V[0-9]S.*', '^V[0-9].*', '.*_V[0-9]', '.*_V[0-9][0-9]',
                            '.*V[0-9]P.*', '.*V[0-9]V.*', '.*[0-9]V[0-9].*', '^[0-9]V.*', '^[0-9][0-9]V.*',
                            '.*[0-9]\.[0-9]V.*', '.*[0-9]_[0-9]V.*', '.*_[0-9]V.*', '.*_[0-9][0-9]V.*',
                            '.*_[0-9]\.[0-9]V.*', '.*[0-9]P[0-9]V.*', '.[0-9]*P[0-9][0-9]V.*', '.*V_[0-9]P[0-9].*',
                            '.*\+[0-9]V.*', '.*\+[0-9][0-9]V.*']
    PWR_Net_List = [net for net in netlist for keyword in PWR_Net_KeyWord_List if re.findall(keyword, net) != []]
    # print(PWR_Net_List)
    PWR_Net_List = sorted(list(set(PWR_Net_List)))

    GND_Net_List = [net for net in netlist if net.find('GND') > -1]
    GND_Net_List = sorted(list(set(GND_Net_List)))

    # 被排除的线：地线和电源线
    Exclude_Net_List = sorted(list(set(PWR_Net_List + GND_Net_List)))

    return Exclude_Net_List, PWR_Net_List, GND_Net_List


# 区分差分，单根与非信号线
def net_separate(diff_pair_brd_data, All_Net_List):

    diff_pair_dict, diff_no_brackets_list = diff_detect(diff_pair_brd_data)

    # 提取差分对
    diff_list = []
    for x in diff_pair_dict.keys():
        diff_list.append([x, diff_pair_dict[x]])

    temp1 = All_Net_List
    ################################################################
    # temp1中存放的是所有信号，temp2中存放的是带N的信号，temp3中存放的是其他信号
    temp2, temp3, temp4 = [], [], []
    for i in range(0, len(temp1)):
        # N只有一种的情况
        if temp1[i].count("N") == 1:
            temp2.append(temp1[i])
        elif temp1[i].count("N") > 1:
            temp3.append(temp1[i])

    # 判断是否是差分对
    # 一个N时
    for i in range(len(temp2)):
        trans2_data = temp2[i].replace("N", "P")
        if trans2_data in temp1:
            temp4.append([temp2[i], trans2_data])

    # 多个N时判断最后一个
    for i in range(len(temp3)):
        last_ind = find_last(temp3[i], 'N')
        list3 = list(temp3[i])
        list3[last_ind] = 'P'
        trans3_data = ''.join(list3)

        if trans3_data in temp1:
            temp4.append([temp3[i], trans3_data])

    # 防止自动生成的差分报告不准确
    for i in temp4:
        if [i[0]] in diff_no_brackets_list:
            # print([i[0]])
            # print(i)
            diff_list.append(i)
            diff_pair_dict[i[0]] = i[1]

        elif [i[1]] in diff_no_brackets_list:
            diff_list.append(i)
            # print([i[1]])
    # print(diff_pair_dict)
    # Get the diff list
    diff_list = sorted(diff_list)

    # print(diff_list)

    # Get the "non_signal_net_list"
    non_signal_net_list, _, _ = get_exclude_netlist(All_Net_List)
    non_signal_net_list = list(set(All_Net_List) & set(non_signal_net_list))
    non_signal_net_list = sorted(non_signal_net_list)

    # Get the single_ended_list
    single_ended_half_list = set(All_Net_List) - set(_flatten(diff_list))
    single_ended_list = list(set(single_ended_half_list) - set(non_signal_net_list))
    single_ended_list = sorted(single_ended_list)

    # print(1111, diff_list)
    # print(2222, single_ended_list)
    # print(3333, non_signal_net_list)
    return diff_pair_dict, diff_list, single_ended_list, non_signal_net_list


"""各种detect功能集合"""


# 对SCH_data进行类化
class SchClass(object):
    def __init__(self, name, model_info, pin_, etch_, pin_net_, pinx_, piny_, coNNsch_):
        self.name = name
        self.model_info = model_info
        self.pin_ = tuple(sorted(list(pin_)))
        self.etch_ = etch_
        self.pin_net_ = pin_net_
        self.pinx_ = pinx_
        self.piny_ = piny_
        self.coNNsch_ = coNNsch_

    def get_name(self):
        return self.name

    def get_model(self):
        return self.model_info

    def get_pin_list(self):
        return self.pin_

    def get_net(self, pin_n):
        return self.pin_net_.get(pin_n)

    def get_net_list(self):
        return self.etch_

    def get_xpoint(self, pin_n):
        return self.pinx_[pin_n]

    def get_ypoint(self, pin_n):
        return self.piny_[pin_n]

    def get_xy(self, pin_n):
        return (self.pinx_[pin_n], self.piny_[pin_n])


# 对net_data进行类化
class NetClass(object):
    def __init__(self, etch_, seg_set, segwidth, seglen, seglayer, segp1, segp2, conn_schpin):
        self.etch_ = etch_
        self.seg_set = seg_set
        self.segwidth = segwidth
        self.seglen = seglen
        self.seglayer = seglayer
        self.segp1 = segp1
        self.segp2 = segp2
        self.conn_schpin = conn_schpin

    def get_name(self):
        return self.etch_

    def get_segment_list(self):
        return self.seg_set

    def get_width(self, seg_id):
        return self.segwidth[seg_id]

    def get_length(self, seg_id):
        return self.seglen[seg_id]

    def get_layer(self, seg_id):
        return self.seglayer[seg_id]

    def get_xy1(self, seg_id):
        return self.segp1[seg_id]

    def get_xy2(self, seg_id):
        return self.segp2[seg_id]

    def get_connected_sch_list(self, seg_id):
        output_sch = list()
        # {('R2', '1'): [[('1710.14', '1076.50'), (1, 2)]], ('U1', 'Y8'): [[('1673.68', '1124.77'), (2, 2)]]}
        for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
            for idx in range(len(pin_seg_info)):
                pinpoint = pin_seg_info[idx][0]
                seg_ind = pin_seg_info[idx][1]
                if seg_ind and seg_ind[0] == seg_id:
                    output_sch.append((SCH_name, pin_id, pinpoint, seg_ind))
                    # eg: ('R2', '1', ('1710.14', '1076.50'), (1, 2))
        return output_sch

    # 根据坐标点是否是线所连pin脚点来判断
    def get_conn_comp(self, seg_ind_input):
        for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
            for idx in range(len(pin_seg_info)):
                seg_ind = pin_seg_info[idx][1]
                if seg_ind and seg_ind == seg_ind_input:
                    return (SCH_name, pin_id)
        return None

    # 返回与芯片相连的信号线的芯片名称与pin脚id
    def get_connected_sch_list_by_seg_ind(self, seg_ind_input):
        connected_SCH_list = list()
        count = 0
        for (SCH_name, pin_id), pin_seg_info in self.conn_schpin.items():
            count += 1
            for idx in range(len(pin_seg_info)):
                seg_ind = pin_seg_info[idx][1]
                if seg_ind and seg_ind == seg_ind_input:
                    connected_SCH_list.append((SCH_name, pin_id))
        if connected_SCH_list != []:
            return connected_SCH_list
        else:
            return None


# 对SCH_brd_data信号报告做一个处理，返回一个对象列表
def SCH_detect(SCH_brd_data, non_signal_net_list):
    # Extract schematic info. from report file

    SCH_object_list = list()
    SCH_Data = SCH_brd_data.GetData()

    SCH_Name_list = list()
    SCH_dict_tmp = dict()
    # print(SCH_Data)
    for line in SCH_Data:
        if line[0]:
            SCH_Name_list.append(line[0])
            # 出现多次情况多次输入
            # get(key, default = None):如果键不存在，返回默认值
            if not SCH_dict_tmp.get(line[0], None):
                # print(line[0])
                # print(line[0], line[1], line[-1], line[3], line[-3], line[-2])
                SCH_dict_tmp[line[0]] = [(line[1], line[-1], line[3], line[-3], line[-2])]
            else:
                SCH_dict_tmp[line[0]].append((line[1], line[-1], line[3], line[-3], line[-2]))

    SCH_Name_list = set(SCH_Name_list)
    # print(SCH_Name_list)
    # print(SCH_dict_tmp)
    for x in SCH_Name_list:
        pin_list = []
        net_list = []
        model_description = []
        x_point_list = []
        y_point_list = []
        for xttt in SCH_dict_tmp[x]:
            pin_list.append(xttt[0])
            # 可能为空，表示pin脚无连线
            net_list.append(xttt[1])
            # 可能为空，描述为空（不确定）
            model_description.append(xttt[2])
            x_point_list.append(xttt[3])
            y_point_list.append(xttt[4])

        pin_net_dict = dict()
        pin_x_dict = dict()
        pin_y_dict = dict()
        for idx_ppp in range(len(pin_list)):
            pin_net_dict[pin_list[idx_ppp]] = net_list[idx_ppp]
            pin_x_dict[pin_list[idx_ppp]] = x_point_list[idx_ppp]
            pin_y_dict[pin_list[idx_ppp]] = y_point_list[idx_ppp]

        # 留下 diff 与 se 信号
        # In order to check connected SCH for any SCH, any net in "Exclude_Net_List" should be excluded
        # check_SCH_net_list也包含空值
        check_SCH_net_list = set(net_list) - set(non_signal_net_list)
        # print(check_SCH_net_list)

        # connected_SCH_list does not include the case that 0ohm resistor connected btw two SCH
        connected_SCH_list = []
        for d in SCH_Data:
            if d[-1] in check_SCH_net_list:  # and d[0] not in [x, '', None]:
                connected_SCH_list.append(d[0])
        # print(2)
        # print(connected_SCH_list)

        # x: 所有元件名  model_description: 描述名   pin_list: 元件个数列表   net_list: 线名列表
        # pin_x_dict: 线的x轴坐标值   pin_y_dict: 线的y轴坐标值  connected_SCH_list: 去除其他信号后的线名
        # （只有diff与se），可能为空
        SCH_object_list.append(SchClass(x, model_description[0], pin_list, net_list, pin_net_dict, pin_x_dict,
                                        pin_y_dict, connected_SCH_list))
    return SCH_object_list


# 对Net_brd_data报告做处理，与SCH_object_list信号对象报告中的pin脚坐标值比较0
# 即找出信号线与pin脚相连的坐标点
def net_detect(Net_brd_data, SCH_object_list, non_signal_net_list):
    # Pre-processing the report file to get the etch_line list

    net_data = Net_brd_data.GetData()
    net_object_list = list()

    # 去除非信号后的线名
    # 这个check_net_list与SCH_detect中的check_net_list不同
    # 这个check_net_list是完整的所有的线名，而SCH_detect中的check_net_list只是所有与pin脚相接的线名
    # 并且这里的线名不会为空
    check_net_list = list(set([x[0] for x in net_data]) - set(non_signal_net_list))

    for net in check_net_list:
        # print (net)
        # 按 check_net_list 中的 net 名顺序列出每个 net 对应的叠层，长宽及过孔的坐标值
        layer_list, width_list, length_list, xy1_list, xy2_list = zip(
            *[(x[2], float(x[5]), float(x[6]), x[-1][0], x[-1][1]) for x in net_data if net == x[0]])
        segment_list = range(1, len(layer_list) + 1)
        # print (net)
        # print(layer_list, width_list, length_list, xy1_list, xy2_list)
        seg_width_dict = dict()
        seg_length_dict = dict()
        seg_layer_dict = dict()
        seg_xy_dict1 = dict()
        seg_xy_dict2 = dict()
        # for id_ss in range(len(segment_list)):
        for id_ss in range(len(layer_list)):
            # print(segment_list[id_ss])
            seg_width_dict[segment_list[id_ss]] = width_list[id_ss]
            seg_length_dict[segment_list[id_ss]] = length_list[id_ss]
            seg_layer_dict[segment_list[id_ss]] = layer_list[id_ss]
            seg_xy_dict1[segment_list[id_ss]] = xy1_list[id_ss]
            seg_xy_dict2[segment_list[id_ss]] = xy2_list[id_ss]

        # Get the connected SCH and Pin (x, y) information to help topology construction
        connected_SCH_Pin_dict_temp = dict()
        for sch in SCH_object_list:
            if net in sch.get_net_list():
                # print(sch.GetNetList())
                for pin in sch.get_pin_list():
                    # print (pin)
                    # 按 check_net_list 中的 net 名顺序排序
                    if sch.get_net(pin) == net:
                        # if net == 'PCH_HSOP0':
                        #     print(sch.GetName())
                        # print (pin + '\n')
                        connected_SCH_Pin_dict_temp[(sch.get_name(), pin)] = (
                        (sch.get_xpoint(pin), sch.get_ypoint(pin)), '')

        connected_SCH_Pin_dict = dict()

        # Determine which segment has SCH connected

        # 存在的意义是什么，不懂
        min_pin_net_connect_distance = 10  # unit = mils

        for (sch, pin), (pin_xy_point, sch_layer) in connected_SCH_Pin_dict_temp.items():

            connected_SCH_Pin_dict[(sch, pin)] = list()

            d1_list = []
            d2_list = []
            for seg in segment_list:
                # SCH_data 中点的坐标值 与 net_data 坐标值的距离
                # print (net, seg)
                # print (pin_xy_point)
                # print (seg_xy_dict1[seg])
                # print (seg_xy_dict2[seg])
                # print ('\n')

                d1_list.append(two_point_distance(pin_xy_point, seg_xy_dict1[seg]))
                d2_list.append(two_point_distance(pin_xy_point, seg_xy_dict2[seg]))

            # print (d1_list)
            d_min = min(d1_list + d2_list + [min_pin_net_connect_distance])
            # print(sch,pin)
            # print(d_min)
            # print (d_min) # 大部分为0，小部分距离在10以内，大于十的去除
            # d_min为0表示线与pin脚相连接
            ind = -1
            # 拿到每条信号线上pin的坐标点
            for idx_d12 in range(len(d1_list)):
                ind += 1
                # 找出与pin脚相接的线
                if d1_list[idx_d12] == d_min:
                    connected_SCH_Pin_dict[(sch, pin)] += [[pin_xy_point, (ind + 1, 1)]]

                    # if net == 'PCH_HSOP0':
                    #     print('01', connected_SCH_Pin_dict)
                elif d2_list[idx_d12] == d_min:
                    connected_SCH_Pin_dict[(sch, pin)] += [[pin_xy_point, (ind + 1, 2)]]
                    # if net == 'PCH_HSOP0':
                    #     print('02', connected_SCH_Pin_dict)

            if connected_SCH_Pin_dict[(sch, pin)] == []:
                connected_SCH_Pin_dict[(sch, pin)] += [[pin_xy_point, None]]
            # print(connected_SCH_Pin_dict)

        # 信号名，每种信号名的range，宽度字典，长度字典，叠层字典，不间断线信号的起始坐标值，经过所有pin脚的坐标值
        net_object_list.append(
            NetClass(net, segment_list, seg_width_dict, seg_length_dict, seg_layer_dict, seg_xy_dict1, seg_xy_dict2,
                     connected_SCH_Pin_dict))

    return net_object_list


"""各种判断name方法的集合"""


# 判断信号名是否在net_data中，并通过net名返回net_object
def get_net_object_by_name(net_name, net_object_list):
    # Get the net object from the net name
    find = False
    for net in net_object_list:
        if net.get_name() == net_name:
            return net
    if ~find:
        # print(net.get_name())
        # print(net_name)
        return None


# 判断信号名是否在SCH_data中，并通过net名返回SCH_object
def get_SCH_object_by_name(SCH_name, SCH_object_list):
    # Get the sch_object from the component name

    for sch in SCH_object_list:
        if sch.get_name() == SCH_name:
            return sch


# 找出与给定芯片连接的信号线的名称
def get_connected_net_list_by_SCH_name(SCH_name, SCH_object_list, net_object_list, non_signal_net_list):
    # Get the connected net list of specified component by name
    # SCH_object_list包含芯片名称，使用的pin脚个数和名称，pin脚连接的线名与坐标值

    check_net_list = []
    for sch in SCH_object_list:
        # print(sch.get_name(), SCH_name)
        if sch.get_name() == SCH_name:
            check_net_list = sch.get_net_list()
            break

    check_net_list_output = []
    for x in check_net_list:
        if x not in non_signal_net_list and len(x) > 0 and get_net_object_by_name(x, net_object_list):
            check_net_list_output.append(x)

    return check_net_list_output


# 取 SCH_brd_data 中每个元素最后一个值，线名
def getallnetlist(SCH_brd_data):
    SCH_Data = SCH_brd_data.GetData()
    All_Net_List = list(set([x[-1] for x in SCH_Data if x[-1] not in ['', None]]))
    return All_Net_List


# 区分 type_list 和 layer_list
def getalllayerlist(stackup_brd_data):
    # Get Layer List from report file

    data = stackup_brd_data.GetData()
    data = data[0:-1]
    for idx in range(len(data)):
        data[idx] = data[idx].split(',')
    layer_list = list()
    type_list = list()

    for layer in data:
        try:
            if layer[1] in ['CONDUCTOR', 'PLANE']:
                layer_list.append(layer[0])
                type_list.append(layer[1])
        except:
            pass

    return layer_list, type_list


# 计算两点距离
def two_point_distance(xy1, xy2):
    x1, y1 = xy1[0], xy1[1]
    x2, y2 = xy2[0], xy2[1]

    d = ((float(x1) - float(x2))**2 + (float(y1) - float(y2))**2)**0.5
    return d


def getpinnumio():
    pin_number_in_out_dict = {('RESA_8P4R_8P4R0402V_33_5%', '8'): ['1'], ('2N7002DWA_7_6PIN_SOT363M_2N7002', 'D2')
    : ['S2'],('2N7002DWA_7_6PIN_SOT363M_2N7002', 'S2'): ['D2'], ('COMMON CHOKE_0805_4_CMC_90_330M', '2'): ['1'],
     ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '11'): ['20'], ('CMC_TYPE1_0805_4_CMC_90_220MA', '4'): ['3'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '35'): ['11'], ('H5143NL_CKS24HC_350UH', '5'): ['20'],
     ('TPS2543_16PIN_QFN_16_20_118X118', '3'): ['10'], ('COMMON CHOKE_1_L4S12X20V_CMC_90', '1'): ['4'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '14'): ['44'], ('2N7002DWA_7_6PIN_SOT363M_DMN65D', 'S1'): ['D1'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '37'): ['9'], ('2N7002DWA_7_6PIN_SOT363M_BSS138', 'S2'): ['D2'],
     ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '7'): ['2'], ('TPS2543_16PIN_QFN_16_20_118X118', '10'): ['3'],
     ('RESA_4P2R_4P2R0402_33_5%', '1'): ['4'], ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '11'): ['28'],
     ('CMC_TYPE1_L4S12X20U_H39_CMC_120', '3'): ['4'], ('COMMON CHOKE_0805_4_CMC_90_330M', '3'): ['4'],
     ('COMMON CHOKE_0805_4_CMC_90_400M', '2'): ['1'], ('H5143NL_CKS24HC_350UH', '22'): ['3'],
     ('COMMON CHOKE_L4S12X20U_CMF_67_5', '4'): ['3'], ('CMC_TYPE1_0805_4_CMC_90_220MA', '3'): ['4'],
     ('RESA_4P2R_4P2R0402_1K_5%', '1'): ['4'], ('TPS2543_16PIN_TQFN_16_20_118X11', '3'): ['10'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '7'): ['32'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '29'): ['8'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '38'): ['8'], ('PS8331B_4_QFN_60_16_197X354_TH1', '42'): ['5'],
     ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '23'): ['8'], ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '20'): ['11'],
     ('RESA_4P2R_4P2R0402_4.7K_5%', '4'): ['1'], ('COMMON CHOKE_L4S12X20U_H39_CMC_', '2'): ['1'],
     ('COMMON CHOKE_0805_4_CMC_90_330M', '4'): ['3'], ('COMMON CHOKE_0805_4_CMC_90_400M', '3'): ['4'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '16'): ['18'], ('H5143NL_CKS24HC_350UH', '2'): ['23'],
     ('CMC_TYPE1_0805_4_CMC_90_220MA', '2'): ['1'], ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '11'): ['20'],
     ('RESA_4P2R_4P2R0402_1K_5%', '2'): ['3'], ('XTAL-SMD-4P_X4S38X80_32.768KHZ', '1'): ['4'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '29'): ['31'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '40'): ['6'],
     ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '5'): ['4'], ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '14'): ['44'],
     ('RESA_8P4R_8P4R0402V_H14_33_5%', '7'): ['2'], ('RESA_4P2R_4P2R0402_39_5%', '1'): ['4'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '26'): ['11'], ('CMC_TYPE1_L4S12X20U_H39_CMC_120', '1'): ['2'],
     ('RESA_4P2R_4P2R0402_0_5%', '3'): ['2'], ('COMMON CHOKE_L4S12X20U_CMF_67_5', '2'): ['1'],
     ('CMC_TYPE1_0805_4_CMC_90_220MA', '1'): ['2'], ('COMMON CHOKE_L4S12X20U_CMC_67_2', '4'): ['3'],
     ('RESA_4P2R_4P2R0402_1K_5%', '3'): ['2'], ('H5143NL_CKS24HC_350UH', '17'): ['8'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '5'): ['34'], ('"COMMON CHOKE_L4S12X20U_CMC_67_2', 'O40X20"'): ['4'],
     ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '8'): ['1'], ('SLG55596AVTR_9PIN_TDFN_8_20_79_', '3'): ['6'],
     ('RESA_8P4R_8P4R0402V_H14_33_5%', '6'): ['3'], ('RESA_4P2R_4P2R0402_39_5%', '2'): ['3'],
     ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '8'): ['23'], ('COMMON CHOKE_L4S12X20U_CMC_120_', '2'): ['1'],
     ('CMC_TYPE1_L4S12X20U_H39_CMC_67_', '4'): ['3'], ('RESA_4P2R_4P2R0402_4.7K_5%', '2'): ['3'],
     ('USB3_REDRIVER_0_QFN24C1_PTN3624', '20'): ['11'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '11'): ['8'],
     ('COMMON CHOKE_0805_4_CMC_90_400M', '1'): ['2'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '23'): ['44'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '9'): ['37'], ('USB3_REDRIVER_0_QFN24C1_PTN3624', '23'): ['8'],
     ('CMC_TYPE1_0805_4_CMC_90_400MA', '2'): ['1'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '31'): ['6'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '23'): ['14'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '15'): ['18'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '11'): ['46'], ('COMMON CHOKE_L4S12X20U_CMC_67_3', '4'): ['3'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '45'): ['15'], ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '19'): ['12'],
     ('2N7002DWA_7_6PIN_SOT363M_BSS138', 'D1'): ['S1'], ('RESA_4P2R_4P2R0402_39_5%', '3'): ['2'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '18'): ['3'], ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '42'): ['14'],
     ('', ''): [''], ('PS8331B_4_QFN_60_16_197X354_TH1', '19'): ['37'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '22'): ['31'], ('RESA_4P2R_4P2R0402_22_5%', '4'): ['1'],
     ('COMMON CHOKE_L4S12X20U_CMC_67_2', '2'): ['1'], ('2N7002DWA_7_6PIN_SOT363M_2N7002', 'S1'): ['D1'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '10'): ['29'], ('H5143NL_CKS24HC_350UH', '13'): ['12'],
     ('COMMON CHOKE_L4S12X20U_H47_CMF_', '2'): ['1'], ('RESA_8P4R_8P4R0402V_H14_33_5%', '4'): ['5'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '20'): ['14'], ('USB3_REDRIVER_0_QFN24C1_PTN3624', '11'): ['20'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '46'): ['1'], ('PS8331B_4_QFN_60_16_197X354_TH1', '12'): ['45'],
     ('2N7002DWA_7_6PIN_SOT363M_DMN65D', 'S2'): ['D2'], ('H5143NL_CKS24HC_350UH', '9'): ['16'],
     ('RESA_4P2R_4P2R0402_0_5%', '4'): ['1'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '25'): ['9'],
     ('H5143NL_CKS24HC_350UH', '6'): ['19'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '12'): ['9'],
     ('TPS2543_16PIN_QFN_16_20_118X118', '11'): ['2'], ('COMMON CHOKE_1_L4S12X20V_CMC_90', '2'): ['3'],
     ('2N7002DWA_7_6PIN_SOT363M_DMN65D', 'D2'): ['S2'], ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '9'): ['22'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '14'): ['43'], ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '2'): ['7'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '6'): ['40'], ('H5143NL_CKS24HC_350UH', '14'): ['11'],
     ('COMMON CHOKE_L4S12X20U_H39_CMC_', '3'): ['4'], ('COMMON CHOKE_L4S12X20U_CMC_67_4', '1'): ['2'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '13'): ['26'], ('PS8331B_4_QFN_60_16_197X354_TH1', '39'): ['7'],
     ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '22'): ['9'], ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '9'): ['22'],
     ('COMMON CHOKE_0805_4_CMC_90_400M', '4'): ['3'], ('H5143NL_CKS24HC_350UH', '16'): ['9'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '37'): ['9'], ('PS8331B_4_QFN_60_16_197X354_TH1', '45'): ['2'],
     ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '22'): ['9'], ('SLG55596AVTR_9PIN_TDFN_8_20_79_', '2'): ['7'],
     ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '12'): ['19'], ('RESA_4P2R_4P2R0402_0_5%', '2'): ['3'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '3'): ['11'], ('COMMON CHOKE_L4S12X20U_CMC_67_3', '1'): ['2'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '27'): ['31'], ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '22'): ['9'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '25'): ['14'], ('RESA_8P4R_8P4R0402V_H14_33_5%', '2'): ['7'],
     ('COMMON CHOKE_L4S12X20U_CMC_67_4', '2'): ['1'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '31'): ['33'],
     ('COMMON CHOKE_L4S12X20U_H39_CMC_', '1'): ['2'], ('PS8331B_4_QFN_60_16_197X354_TH1', '7'): ['39'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '41'): ['4'], ('H5143NL_CKS24HC_350UH', '12'): ['13'],
     ('1', ('1', '2')): ['2'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '27'): ['10'],
     ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '9'): ['22'], ('PS8331B_4_QFN_60_16_197X354_TH1', '17'): ['39'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '40'): ['6'], ('RESA_8P4R_8P4R0402V_H14_33_5%', '1'): ['8'],
     ('COMMON CHOKE_L4S12X20U_CMC_67_4', '3'): ['4'], ('CMC_TYPE1_L4S12X20U_H39_CMC_120', '2'): ['1'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '4'): ['43'], ('USB3_REDRIVER_0_QFN24C1_PTN3624', '8'): ['23'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '32'): ['7'], ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '26'): ['13'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '32'): ['34'], ('TPS2543_16PIN_TQFN_16_20_118X11', '2'): ['11'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '42'): ['3'], ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '19'): ['12'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '14'): ['25'], ('PS8331B_4_QFN_60_16_197X354_TH1', '30'): ['32'],
     ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '8'): ['23'], ('COMMON CHOKE_L4S12X20U_CMC_67_3', '3'): ['4'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '17'): ['2'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '30'): ['7'],
     ('COMMON CHOKE_L4S12X20U_CMC_67_4', '4'): ['3'], ('COMMON CHOKE_L4S12X20U_CMC_120_', '4'): ['3'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '5'): ['42'], ('USB3_REDRIVER_0_QFN24C1_PTN3624', '9'): ['22'],
     ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '3'): ['11'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '44'): ['14'],
     ('2', ('1', '2')): ['1'], ('RESA_4P2R_4P2R0402_22_5%', '1'): ['4'],
     ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '41'): ['15'], ('CMC_TYPE1_0805_4_CMC_90_400MA', '4'): ['3'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '36'): ['10'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '24'): ['13'],
     ('SLG55596AVTR_9PIN_TDFN_8_20_79_', '7'): ['2'], ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '31'): ['8'],
     ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '6'): ['3'], ('RESA_8P4R_8P4R0402V_33_5%', '3'): ['6'],
     ('COMMON CHOKE_L4S12X20U_CMC_67_3', '2'): ['1'], ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '9'): ['12'],
     ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '19'): ['12'], ('PS8331B_4_QFN_60_16_197X354_TH1', '2'): ['45'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '15'): ['42'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '21'): ['42'],
     ('RESA_8P4R_8P4R0402V_H14_33_5%', '8'): ['1'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '33'): ['3'],
     ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '8'): ['11'], ('RESA_4P2R_4P2R0402_22_5%', '2'): ['3'],
     ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '12'): ['19'], ('SLG55596AVTR_9PIN_TDFN_8_20_79_', '6'): ['3'],
     ('RESA_8P4R_8P4R0402V_H14_33_5%', '5'): ['4'], ('RESA_8P4R_8P4R0402V_33_5%', '2'): ['7'],
     ('COMMON CHOKE_L4S12X20U_H47_CMF_', '4'): ['3'], ('RESA_4P2R_4P2R0402_33_5%', '3'): ['2'],
     ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '11'): ['8'], ('PI2DBS212ZHE_2_QFN_28_20_138X21', '3'): ['20', '5'],
     ('H5143NL_CKS24HC_350UH', '20'): ['5'], ('H5143NL_CKS24HC_350UH', '3'): ['22'],
     ('COMMON CHOKE_L4S12X20U_H39_CMC_', '4'): ['3'], ('CMC_TYPE1_0805_4_CMC_90_400MA', '3'): ['4'],
     ('RESA_4P2R_4P2R0402_22_5%', '3'): ['2'], ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '45'): ['15'],
     ('COMMON CHOKE_1_L4S12X20V_CMC_90', '4'): ['1'], ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '23'): ['8'],
     ('TPS2543_16PIN_TQFN_16_20_118X11', '10'): ['3'], ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '4'): ['5'],
     ('RESA_8P4R_8P4R0402V_33_5%', '5'): ['4'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '36'): ['10'],
     ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '15'): ['45'], ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '11'): ['20'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '29'): ['10'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '42'): ['14'],
     ('USB3_REDRIVER_0_QFN24C1_PTN3624', '19'): ['12'], ('PI2DBS212ZHE_2_QFN_28_20_138X21', '2'): ['21', '4'],
     ('COMMON CHOKE_L4S12X20U_CMF_67_5', '3'): ['4'], ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '23'): ['8'],
     ('H5143NL_CKS24HC_350UH', '8'): ['17'], ('XTAL_25.000M_SMD-4P_X4S32X25_H3', '3'): ['1'],
     ('TPS2543_16PIN_TQFN_16_20_118X11', '11'): ['2'], ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '4'): ['35'],
     ('H5143NL_CKS24HC_350UH', '23'): ['2'], ('XTAL-48M-SMD-4P_X4S32X25_H31_48', '3'): ['1'],
     ('COMMON CHOKE_L4S12X20U_H47_CMF_', '3'): ['4'], ('RESA_8P4R_8P4R0402V_33_5%', '4'): ['5'],
     ('RESA_4P2R_4P2R0402_33_5%', '4'): ['1'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '26'): ['8'],
     ('COMMON CHOKE_L4S12X20U_CMC_120_', '3'): ['4'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '28'): ['9'],
     ('RESA_4P2R_4P2R0402_4.7K_5%', '3'): ['2'], ('2N7002DWA_7_6PIN_SOT363M_BSS138', 'S1'): ['D1'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '1'): ['46'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '39'): ['7'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '34'): ['5'], ('RESA_4P2R_4P2R0402_1K_5%', '4'): ['1'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '41'): ['15'], ('CMC_TYPE1_0805_4_CMC_90_400MA', '1'): ['2'],
     ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '12'): ['19'], ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '2'): ['12'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '15'): ['45'], ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '28'): ['11'],
     ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '3'): ['6'], ('PS8331B_4_QFN_60_16_197X354_TH1', '24'): ['31'],
     ('PI3EQX7502AIZDEX_24PIN_2_QFN_24', '20'): ['11'], ('RESA_8P4R_8P4R0402V_33_5%', '7'): ['2'],
     ('CMC_TYPE1_L4S12X20U_H39_CMC_67_', '3'): ['4'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '33'): ['31'],
     ('XTAL_25.000M_SMD-4P_X4S32X25__1', '3'): ['1'], ('PS8331B_4_QFN_60_16_197X354_TH1', '43'): ['4'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '16'): ['40'], ('CMC_TYPE1_L4S12X20U_H39_CMC_120', '4'): ['3'],
     ('USB3_REDRIVER_0_QFN24C1_PTN3624', '22'): ['9'], ('TPS2543_16PIN_QFN_16_20_118X118', '2'): ['11'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '31'): ['24'], ('XTAL_25.000M_SMD-4P_X4S32X25_H3', '1'): ['3'],
     ('COMMON CHOKE_L4S12X20U_CMF_67_5', '1'): ['2'], ('USB3_REDRIVER_0_QFN24C1_PTN3624', '12'): ['19'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '8'): ['11'], ('COMMON CHOKE_L4S12X20U_CMC_67_2', '3'): ['4'],
     ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '2'): ['12'], ('H5143NL_CKS24HC_350UH', '19'): ['6'],
     ('RESA_4P2R_4P2R0402_39_5%', '4'): ['1'], ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '44'): ['14'],
     ('COMMON CHOKE_L4S12X20U_H47_CMF_', '1'): ['2'], ('RESA_8P4R_8P4R0402V_33_5%', '6'): ['3'],
     ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '19'): ['13'], ('CMC_TYPE1_L4S12X20U_H39_CMC_67_', '2'): ['1'],
     ('COMMON CHOKE_L4S12X20U_CMC_120_', '1'): ['2'], ('XTAL_25.000M_SMD-4P_X4S32X25__1', '1'): ['3'],
     ('RESA_4P2R_4P2R0402_4.7K_5%', '1'): ['4'], ('PI2DBS212ZHE_2_QFN_28_20_138X21', '7'): ['16', '9'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '20'): ['36'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '34'): ['32'],
     ('PS8803CQFN52GTR-A1_52PIN_QFN_52', '12'): ['9'], ('COMMON CHOKE_1_L4S12X20V_CMC_90', '3'): ['2'],
     ('2N7002DWA_7_6PIN_SOT363M_2N7002', 'D1'): ['S1'], ('XTAL-48M-SMD-4P_X4S32X25_H31_48', '1'): ['3'],
     ('XTAL-SMD-4P_X4S38X80_32.768KHZ', '4'): ['1'], ('RESA_8P4R_8P4R0402V_H18_4.7K_5%', '1'): ['8'],
     ('RESA_8P4R_8P4R0402V_H14_33_5%', '3'): ['6'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '20'): ['41'],
     ('RESA_4P2R_4P2R0402_0_5%', '1'): ['4'], ('SN65LVPE501_1_QFN24C1_SN65LVPE5', '20'): ['11'],
     ('RESA_8P4R_8P4R0402V_33_5%', '1'): ['8'], ('PI3WVR12412ZHE_42PIN_TQFN42C_PI', '32'): ['4'],
     ('CMC_TYPE1_L4S12X20U_H39_CMC_67_', '1'): ['2'], ('PS8331B_4_QFN_60_16_197X354_TH1', '32'): ['25'],
     ('COMMON CHOKE_0805_4_CMC_90_330M', '1'): ['2'], ('PS8331B_4_QFN_60_16_197X354_TH1', '10'): ['36'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '35'): ['4'], ('SN65LVPE501_0_QFN24C1_SN65LVPE5', '8'): ['23'],
     ('PI2DBS212ZHE_2_QFN_28_20_138X21', '6'): ['17', '8'], ('2N7002DWA_7_6PIN_SOT363M_BSS138', 'D2'): ['S2'],
     ('2N7002DWA_7_6PIN_SOT363M_DMN65D', 'D1'): ['S1'], ('COMMON CHOKE_L4S12X20U_CMC_67_2', '1'): ['2'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '25'): ['32'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '24'): ['45'],
     ('PI3EQX8904_42P_TQFN42C_PI3EQX89', '8'): ['31'], ('PS8331B_4_QFN_60_16_197X354_TH1', '23'): ['32'],
     ('PS8331B_4_QFN_60_16_197X354_TH1', '28'): ['32'], ('H5143NL_CKS24HC_350UH', '11'): ['14'],
     ('RESA_4P2R_4P2R0402_33_5%', '2'): ['3'], ('PS8802QFN52GTR-A0_52PIN_QFN_52_', '9'): ['12']}
    return pin_number_in_out_dict


# 判断是否为浮点数
def isfloat(string):
    try:
        float(string)
        return True
    except Exception:
        return False


'''关于topology的功能模块'''


# 返回与本信号线相连的下一段信号线的名称
def net_mapping(sch_name, input_pin, SCH_object_list, flag = False):
    # Find the input-output mapping of component
    if sch_name is not None:
        sch = get_SCH_object_by_name(sch_name, SCH_object_list)
        model_name = sch.get_model()
        # Two pin component: RES, CAP,...
        pin_number_in_out_dict = getpinnumio()
        output_pin_list = pin_number_in_out_dict.get((input_pin, sch.get_pin_list()))
        # Find output pin by model name
        if output_pin_list is None:
            output_pin_list = pin_number_in_out_dict.get((model_name, input_pin))
        # print(output_pin_list)
        if output_pin_list is None:
            return [None], [None]
        else:
            return [sch.get_net(output_pin) for output_pin in output_pin_list], output_pin_list
    else:
        return [None], [None]


# 列出各段的芯片pin脚，叠层，宽度，长度
# 例如 ['USB3_CMC_TXDN1', '[FRONT_USB_HEADER-5]:BOTTOM:5.1', 976.98, 'TOP:5.1:[U4-2]', 65.81, '[U4-2]:TOP:5.1:[U4-9]',
# 40.0, '[U4-9]:TOP:5.1:[LU2-2]', 81.27]
def topology_format(net_name, topology_seg_ind, net_object_list):
    # Formatting function of exported results

    net = get_net_object_by_name(net_name, net_object_list)
    # if net_name == 'PCH_HSOP0':
    #     print(net.GetName(), net.GetSegmentList(), net.GetWidth(), net.GetLength(), net.GetLayer())
    topology_formatted = list()

    for i in range(len(topology_seg_ind)):
        layer = net.get_layer(topology_seg_ind[i][0])
        width = net.get_width(topology_seg_ind[i][0])
        length = net.get_length(topology_seg_ind[i][0])

        # if net_name == 'USB3_CMC_TXDN1':
        #     print(i,layer, width, length)
        # 不懂为什么要将奇偶情况分开讨论
        if float(length) > 1:
            if i % 2 == 0:
                # 判断是否是芯片相接
                # 与芯片相接则写成 [sch-pin]:layer:width 的形式
                if net.get_connected_sch_list_by_seg_ind(topology_seg_ind[i]) is not None:
                    content = '%s:%s'%(layer, width)
                    for (sch, pin) in net.get_connected_sch_list_by_seg_ind(topology_seg_ind[i]):
                        content = '[%s-%s]:'%(sch, pin) + content
                        # if net_name == 'PCH_HSOP0':
                        #     print(1, content)
                # 如不与芯片相接则写成 layer:width 的形式
                else:
                    content = '%s:%s'%(layer, width)
                # print(1,content)
            else:
                if net.get_connected_sch_list_by_seg_ind(topology_seg_ind[i]) is not None:
                    for (sch, pin) in net.get_connected_sch_list_by_seg_ind(topology_seg_ind[i]):
                        content = content + ':[%s-%s]'%(sch, pin)
                # print(2, content)
                topology_formatted.append(content)
                topology_formatted.append(length)

    return [net_name] + topology_formatted


# 返回每条信号线本身连接芯片Pin脚的标号list（不包含换线）
def topology_extract1(excel_path, net_name, start_sch_name, net_object_list, start_sch_pin=None):
    # Topology Extraction Function
    # 通过net_name返回net_object
    check_net = get_net_object_by_name(net_name, net_object_list)
    # 同一个net_name的线的个数
    seg_number_list = check_net.get_segment_list()
    seg_ind_list_original = list()
    seg_inter_map_dict = dict()
    xy_point_list = list()

    # print(seg_number_list)
    # 每条信号线宽度分割段的数量list,eg:seg = range(1,3)
    for seg in seg_number_list:
        # 每条连续信号线都有起始和终止点坐标所以分1,2

        # ind_list
        seg_ind_list_original.append((seg, 1))
        seg_ind_list_original.append((seg, 2))
        # ind_dict
        seg_inter_map_dict[(seg, 1)] = (seg, 2)
        seg_inter_map_dict[(seg, 2)] = (seg, 1)
        # ind_坐标点
        xy_point_list.append(check_net.get_xy1(seg))
        xy_point_list.append(check_net.get_xy2(seg))

    seg_ind_xy_point_dict = dict()
    for iddx_seg in range(len(seg_ind_list_original)):
        seg_ind_xy_point_dict[seg_ind_list_original[iddx_seg]] = xy_point_list[iddx_seg]

    sch_list = []
    seg_ind_list_original_start = [x for x in seg_ind_list_original]
    # print(seg_ind_list_original_start)
    for seg in seg_number_list:
        # eg: 1 2 3
        # 遍历每条线所连接的器件名
        for sch in check_net.get_connected_sch_list(seg):

            ######################
            sch_list.append(sch)
            ######################
    sch_name_list = [x[0] for x in sch_list]

    topology_seg_ind_list_all = list()

    while start_sch_name in sch_name_list:

        sch_name_list = [x[0] for x in sch_list]
        start_seg_ind1_list = list()
        for sch in sch_list:

            if not start_seg_ind1_list:

                # 信号名相同
                # sch = (SCH_name, pin_id, pinpoint, seg_ind)
                # 找到开始的芯片所连的信号线
                if sch[0] == start_sch_name:

                    # 第一个判断是自己输入起始pin名，第二个判断是默认全部pin名
                    if start_sch_pin is not None and start_sch_pin == sch[1] or start_sch_pin is None:
                        start_seg_ind1_list.append(sch[-1])
                    sch_list.remove(sch)

        if start_seg_ind1_list != [] and seg_ind_list_original == []:
            # print('in')
            # print(seg_ind_list_original_start)
            seg_ind_list_original = seg_ind_list_original_start
        start_seg_ind2_list = []

        for x in start_seg_ind1_list:
            start_seg_ind2_list.append(seg_inter_map_dict[x])

        if start_seg_ind1_list:
            for start_idx in range(len(start_seg_ind1_list)):
                # if net_name == 'USB3_TXDN1_C':
                # print('A', seg_ind_list_original)
                    # for ind in seg_ind_list_original:
                    #     print(seg_ind_xy_point_dict[ind])
                try:
                    seg_ind_list_original.remove(start_seg_ind1_list[start_idx])
                    seg_ind_list_original.remove(start_seg_ind2_list[start_idx])
                except ValueError:
                    pass
                # print('B',seg_ind_list_original)
                end_all = False
                i = 0
                # 不能写成 ~end_all，会出错，不知道为什么 --Gorgeous
                while end_all is False:
                    i += 1
                    # print(i)
                    topology_seg_ind_list = list()
                    # 存入从开始到结束的信号线名称
                    topology_seg_ind_list.append(start_seg_ind1_list[start_idx])
                    topology_seg_ind_list.append(start_seg_ind2_list[start_idx])

                    # 没必要，因为两者无区别
                    seg_ind_list = list(seg_ind_list_original)

                    end = False
                    j = 0
                    while end is False:
                        j += 1
                        # print(j)
                        # 从开始的线依次向后找寻坐标点重合的信号线代表是信号下一个经过的信号线
                        for seg_ind in seg_ind_list:

                            # seg_ind_list中的每一个与topology_seg_ind_list最后一个相比，找出坐标相同点
                            # 判断线是否完整
                            if seg_ind != topology_seg_ind_list[-1] and seg_ind_xy_point_dict[seg_ind] ==\
                                    seg_ind_xy_point_dict[topology_seg_ind_list[-1]]:

                                topology_seg_ind_list.append(seg_ind)
                                topology_seg_ind_list.append(seg_inter_map_dict[seg_ind])
                                # print(2)
                                # print(topology_seg_ind_list[-1])
                                seg_ind_list.remove(seg_ind)
                                seg_ind_list.remove(seg_inter_map_dict[seg_ind])
                                break

                        # Check if any connected segment to the topology_seg_ind_list[-1]
                        end = True

                        # 说明线没找完，继续往下寻找
                        for seg_ind in seg_ind_list:
                            if seg_ind != topology_seg_ind_list[-1] and seg_ind_xy_point_dict[seg_ind] ==\
                                    seg_ind_xy_point_dict[topology_seg_ind_list[-1]]:
                                end = False
                                break

                        # Find the last split point to the end for this line, and remove it from the seg_
                        # ind_list_original
                        split_ind_list = []

                        if end is True:

                            # 找分叉点
                            for ind1 in range(len(topology_seg_ind_list)):
                                for ind2 in range(len(seg_ind_list)):
                                    if topology_seg_ind_list[ind1] != seg_ind_list[ind2] and seg_ind_xy_point_dict[\
                                            topology_seg_ind_list[ind1]] == seg_ind_xy_point_dict[seg_ind_list[ind2]]:
                                        split_ind_list.append(ind1)

                            # Record the split point
                            # 看不懂2代表什么，有什么作用
                            if 0 not in split_ind_list:
                                split_ind_list.append(2)

                            for ind in range(len(topology_seg_ind_list)):
                                if ind >= max(split_ind_list):
                                    if topology_seg_ind_list[ind] not in [start_seg_ind1_list[start_idx],
                                                                          start_seg_ind2_list[start_idx]]:

                                        seg_ind_list_original.remove(topology_seg_ind_list[ind])

                            if split_ind_list == []:
                                end = False

                        # 为什么为50
                        if i > 30 or j > 30:
                            error_message = '%s-%s can\'t be extracted' % (start_sch_name, net_name)
                            create_error_message(excel_path, error_message)
                            raise ValueError('%s-%s can\'t be extracted' % (start_sch_name, net_name))

                    topology_seg_ind_list_all.append(topology_seg_ind_list)

                    end_all = True

                    for seg_ind in seg_ind_list_original:
                        if seg_ind != start_seg_ind2_list[start_idx] and seg_ind_xy_point_dict[seg_ind]\
                                == seg_ind_xy_point_dict[start_seg_ind2_list[start_idx]]:
                            end_all = False
                            break
    if len(seg_ind_list_original) == 0 or start_seg_ind1_list == []:

        connected_sch_list = list()
        end = False
        while end is False:
            end = True
            for line in topology_seg_ind_list_all:
                # print(line)
                for idx in range(len(line)):
                    if len(line) >= 4:
                        # 去除掉第一个与最后一个pin点坐标
                        # print(len(line))
                        if 0 < idx <= len(line)-3:
                            # print(idx)
                            # Construct the tree line
                            # 返回(sch, pin_id)
                            connected_sch_tmp = check_net.get_conn_comp(line[idx])
                            # print(line[idx])
                            # print(connected_sch_tmp)
                            # 如果是pin脚坐标点
                            if connected_sch_tmp not in connected_sch_list + [None]:
                                # 去除最后一段线
                                topology_seg_ind_list_all.append(line[0:idx+1])
                                # print(line[0:idx+1])
                                # 保存从一个pin脚到另一个pin脚中间不经过芯片的同名信号线
                                connected_sch_list.append(connected_sch_tmp)
                                end = False
                                break
                if not end:
                    break

        return topology_seg_ind_list_all
    else:
        error_message = '%s-%s can\'t be extracted' % (start_sch_name, net_name)
        create_error_message(excel_path, error_message)
        raise ValueError('%s-%s can\'t be extracted' % (start_sch_name, net_name))


# 返回每个信号线的经过芯片，叠层，pin脚名，线长
def topology_extract2(start_net_name, start_sch_name, SCH_object_list, net_object_list, non_signal_net_list,
                      excel_path):
    # Topology Extraction Function 2
    # start_time = time.clock()
    # Initialize
    topology_list = list()

    try:
        topology_seg_ind = topology_extract1(excel_path, start_net_name, start_sch_name, net_object_list,
                                             start_sch_pin=None)
    except:
        error_message = '%s-%s can\'t be extracted' % (start_sch_name, net_name)
        create_error_message(excel_path, error_message)
        raise ValueError('%s-%s can\'t be extracted!' % (start_sch_name, start_net_name))

    net_class = get_net_object_by_name(start_net_name, net_object_list)
    topology_return_list = []
    # 显示一根信号线的数据，并指出与之相连的下一根信号线
    for line in topology_seg_ind:
        sch_list, pin_list, net_list, next_pin_list = [], [], [], []

        # 判断最后一个线名是否与芯片相连接,因为经过过孔或者信号线宽度变化也会分段
        # 最后一个线名是否与芯片相连接才能进入循环
        if net_class.get_connected_sch_list_by_seg_ind(line[-1]):

            # Find the connected component of etch line
            # 找出最后一个线名的坐标，其实找出了每条信号线上所经过的所有芯片（pin脚）的坐标值
            if line[-1][1] == 1:
                net_xy_point = net_class.get_xy1(line[-1][0])
            elif line[-1][1] == 2:
                net_xy_point = net_class.get_xy2(line[-1][0])

            # 找出与最后条线相连的芯片的坐标
            # 因此net_xy_point与sch_pin_xy_point_list内坐标其实是相同的
            sch_pin_xy_point_list = list()
            for (sch, pin) in net_class.get_connected_sch_list_by_seg_ind(line[-1]):
                sch_class = get_SCH_object_by_name(sch, SCH_object_list)
                sch_pin_xy_point_list.append((sch_class.get_xpoint(pin), sch_class.get_ypoint(pin)))
            d_xy_list = []

            # 因为两者值相同，所以d_xy_list全为0，考虑是否能省略这部分代码
            for x in sch_pin_xy_point_list:
                d_xy_list.append(two_point_distance(net_xy_point, x))

            ######################################################################################
            # index()方法检测字符创中是否包含字符串str，存在返回索引值，不存在抛出异常
            for ind in range(len(net_class.get_connected_sch_list_by_seg_ind(line[-1]))):
                sch_list.append(net_class.get_connected_sch_list_by_seg_ind(line[-1])[ind][0])
                pin_list.append(net_class.get_connected_sch_list_by_seg_ind(line[-1])[ind][1])

        else:
            # 得出每条信号线所经过的芯片的名称以及pin脚名称列表
            sch_list.append(None)
            pin_list.append(None)

        sch_nochange_list = sch_list

        for ind in range(len(sch_nochange_list)):

            next_net, next_pin = net_mapping(sch_nochange_list[ind], pin_list[ind], SCH_object_list)

            # next_net通常为单个线名，下列代码是否可以改写为 if next_net:
            for idxx_ in range(len(next_net)):
                # start_net_name为每段信号线的名称， line为从两段线开始的ind对
                topology_list.append(topology_format(start_net_name, line, net_object_list))

                net_list.append(next_net[idxx_])
                next_pin_list.append(next_pin[idxx_])

                # 如果存在信号线，则存入信号线所经过的芯片名称
                if idxx_ > 0:
                    sch_list.append(sch_list[-1])
                # 没看懂什么意思
                if net_list[-1] in non_signal_net_list:
                    topology_list[-1].append(net_list[-1])

            net_list_start = net_list

            end = True

            # Ending Condition Detect
            for net in [x for x in net_list if x not in non_signal_net_list]:
                if net is not None:
                    end = False
            j = -1

            pre_net_list = []
            topology_out_list = []

            # Topology Detect for secondary part (net change)
            while end is False:
                # if start_net_name == 'SUSCLK_M2230' and start_sch_name == 'J38':
                #     print(j, end)
                j += 1
                topology_list_temp, net_list_temp, sch_list_temp, pin_list_temp, next_pin_list_temp = [], [], [], [], []

                i = -1
                end = True
                my_flag = 0
                # print(1111, topology_list)
                for ij1 in range(len(net_list)):
                    i += 1
                    # 下段代码为重复代码，可以写成函数简化
                    # 只进入信号线名
                    if net_list[ij1] not in non_signal_net_list+[None]:
                        net_class = get_net_object_by_name(net_list[ij1], net_object_list)

                        pre_net_list.append(net_list[ij1])

                        # if start_net_name == 'PCH_HSON2' and start_sch_name == 'JN44':
                        #     print(j,net_list[ij1])
                        #     print(j,next_pin_list[ij1])
                        #     print(j,sch_list[ij1])

                        try:
                            # 找出下一条线的段id
                            topology_seg_ind = topology_extract1(excel_path, net_list[ij1], sch_list[ij1],
                                                                 net_object_list, start_sch_pin=next_pin_list[ij1])

                            for ind in range(len(topology_seg_ind)):
                                # 找出与芯片相接的线
                                if net_class.get_connected_sch_list_by_seg_ind(topology_seg_ind[ind][-1]) is not None:
                                    # 存入线的两端坐标值
                                    if topology_seg_ind[ind][-1][1] == 1:
                                        net_xy_point = net_class.get_xy1(topology_seg_ind[ind][-1][0])
                                    elif topology_seg_ind[ind][-1][1] == 2:
                                        net_xy_point = net_class.get_xy2(topology_seg_ind[ind][-1][0])
                                    sch_pin_xy_point_list = list()

                                    # if start_net_name == 'PCH_HSOP3':
                                    #     print(net_xy_point)
                                    for (sch, pin) in net_class.get_connected_sch_list_by_seg_ind(
                                            topology_seg_ind[ind][-1]):
                                        sch_class = get_SCH_object_by_name(sch, SCH_object_list)
                                        sch_pin_xy_point_list.append((sch_class.get_xpoint(pin),
                                                                      sch_class.get_ypoint(pin)))
                                    # if start_net_name == 'PCH_HSOP3':
                                    #     print(sch_pin_xy_point_list)
                                    d_xy_list = []
                                    for x in sch_pin_xy_point_list:
                                        d_xy_list.append(two_point_distance(net_xy_point, x))

                                    sch_list_temp.append(net_class.get_connected_sch_list_by_seg_ind(
                                        topology_seg_ind[ind][-1])[d_xy_list.index(max(d_xy_list))][0])
                                    pin_list_temp.append(net_class.get_connected_sch_list_by_seg_ind(
                                        topology_seg_ind[ind][-1])[d_xy_list.index(max(d_xy_list))][1])

                                    # if start_net_name == 'PCH_HSOP3':
                                    #     print(sch_list_temp)
                                    #     print(pin_list_temp)
                                else:
                                    sch_list_temp.append(None)
                                    pin_list_temp.append(None)

                                next_net, next_pin = net_mapping(sch_list_temp[-1], pin_list_temp[-1], SCH_object_list)

                                # if start_net_name == 'PCH_HSOP3':
                                #     print(next_net, next_pin)
                                ########################################################
                                # 自己修改的代码
                                start_flag = 1
                                for idxx_ in range(len(next_net)):
                                    for start_ind in range(len(net_list_start)):
                                        if next_net[idxx_] == net_list_start[start_ind]:
                                            # and next_pin[idxx_] == next_pin_list_start[start_ind]\
                                            # and sch_list_temp[idxx_] == sch_list_start[start_ind]:
                                            start_flag = 0
                                        if start_ind == len(net_list_start) - 1 and start_flag:
                                            net_list_temp.append(next_net[idxx_])
                                            next_pin_list_temp.append(next_pin[idxx_])

                                    ########################################################

                                    if idxx_ > 0:
                                        sch_list_temp.append(sch_list_temp[-1])

                                    topology_list_temp.append(topology_list[i] + topology_format(net_list[ij1],
                                                                            topology_seg_ind[ind], net_object_list))
                                    topology_out_list.append(topology_list[i])
                                    topology_out_list.append(
                                        topology_list[i] + topology_format(net_list[ij1], topology_seg_ind[ind],
                                                                           net_object_list))
                        except:
                            topology_list_temp.append(topology_list[i]+['Exception;%s' % net_list[ij1]])
                            topology_out_list.append(topology_list[i]+['Exception;%s' % net_list[ij1]])
                            sch_list_temp.append(None)
                            pin_list_temp.append(None)
                            net_list_temp.append(None)
                            next_pin_list_temp.append(None)

                    elif net_list[ij1] in non_signal_net_list and j != 0:
                        topology_list_temp.append(topology_list[i] + [net_list[ij1]])
                        topology_out_list.append(topology_list[i] + [net_list[ij1]])
                        sch_list_temp.append(None)
                        pin_list_temp.append(None)
                        net_list_temp.append(None)
                        next_pin_list_temp.append(None)
                    else:
                        topology_list_temp.append(topology_list[i])
                        topology_out_list.append(topology_list[i])
                        # if start_net_name == 'GPP_CLK1N_LAN':
                        #     print(topology_list[j])
                        sch_list_temp.append(None)
                        pin_list_temp.append(None)
                        net_list_temp.append(None)
                        next_pin_list_temp.append(None)
                    # if start_net_name == 'DPC_AUX_DP_C':
                    #     print(topology_list)
                # print(2, topology_list_temp)
                topology_list = list(topology_list_temp)
                # if start_net_name == 'DPC_AUX_DP_C':
                #     print(topology_list)

                net_list = list(net_list_temp)

                next_pin_list = list(next_pin_list_temp)

                sch_list = list(sch_list_temp)

                for net in net_list:
                    if net in pre_net_list:
                        end = True
                        break
                    if net is not None:
                        end = False
                        break
                    if net is None:
                        end = True

                # Unsolvable condition
                if j > 20:
                    error_message = '%s-%s can\'t be extracted' % (start_sch_name, net_name)
                    create_error_message(excel_path, error_message)
                    raise ValueError("!!!!!!!Can't Extract %s-%s!!!!!!!" % (start_sch_name, start_net_name))

            topology_return_half_list = []

            if topology_out_list:
                for x in topology_out_list:
                    if x not in topology_return_half_list:
                        topology_return_half_list.append(x)
            if topology_return_half_list:
                topology_return_list.append(topology_return_half_list)
    if topology_return_list:
        return topology_return_list
    else:
        return [topology_list]


# 简化 topology_extract2 生成的数据格式
# 返回每个信号线的起始芯片和终止芯片，换层次数，总长，加上topology_extract2的数据
def topology_list_format_simplified(topology_list, non_signal_net_list, All_Net_List):

    # Topology Formatting Function
    topology_out_list = []

    signal_net_list = list(set(All_Net_List) ^ set(non_signal_net_list))

    # Calculation Total Length and Via Count for each line
    for idx1 in range(len(topology_list)):
        via_count = 0
        total_length = 0
        net_count = 0
        remove_ind = None
        # print(topology_list)

        if str(topology_list[idx1][-1]).find('Exception;') > -1:
            # print(2, topology_list)
            end_sch_name = topology_list[idx1][-1]
            topology_list[idx1] = [topology_list[idx1][0], 'via_count', 'total_length'] + topology_list[idx1][1::]
        else:
            # 获得信号线最后芯片名并去除[]符号
            # print(0, topology_list[idx1])
            # print(1, topology_list[idx1][-2])
            # print(2, str(topology_list[idx1][-2]).split(':')[-1].find('['))
            if str(topology_list[idx1][-2]).split(':')[-1].find('[') == 0:
                end_sch_name = topology_list[idx1][-2].split(':')[-1][1:-1]

            elif str(topology_list[idx1][-3]).split(':')[-1].find('[') == 0:
                end_sch_name = topology_list[idx1][-3].split(':')[-1][1:-1]
                # print(end_sch_name)
            else: end_sch_name = 'NONE'
            # print(333, end_sch_name)
            for idx2 in range(len(topology_list[idx1])):
                if isfloat(topology_list[idx1][idx2]):
                    total_length += float(topology_list[idx1][idx2])

                if str(topology_list[idx1][idx2]).find(':') > -1:
                    # eg: '[FRONT_USB_HEADER-2]:BOTTOM:5.1'
                    layer = None
                    for x in topology_list[idx1][idx2].split(':'):
                        layer = x

                    layer_next = None
                    for x in topology_list[idx1][idx2+2::]:
                        if str(x).find(':') > -1:
                            for x_ in x.split(':'):
                                layer_next = x_
                                # 换层计数
                                if layer_next != layer:
                                    via_count += 1
                            break
            # via_count 是换层的次数，total_length 是此条信号线加下条信号线的总长度（如果有下条信号线的话）
            topology_list[idx1] = [topology_list[idx1][0], 'via_count %d' % via_count, 'total_length %.3f'
                                   % total_length] + topology_list[idx1][1::]
            # print(2, topology_list)

        if topology_list[idx1][3].find('[') == 0:
            # 获得信号线起始芯片名并去除[]符号
            start_sch_name = topology_list[idx1][3].split(':')[0][1:-1]

        topology_list[idx1] = [start_sch_name, topology_list[idx1][0], end_sch_name] + topology_list[idx1][1::]
        # print(3, topology_list)

        for idx2 in range(len(topology_list[idx1])):
            # 过滤topology_list中的非信号线
            if topology_list[idx1][idx2] in signal_net_list:
                net_count += 1
                if net_count > 1:
                    # 标出信号线
                    topology_list[idx1][idx2] = 'net$%s' % topology_list[idx1][idx2]
            elif topology_list[idx1][idx2] in non_signal_net_list:
                # 删除非信号线
                remove_ind = idx2
        if remove_ind is not None:
            # 其实就是删除GND信号
            topology_list[idx1].pop(remove_ind)

        topology_out_list.append(topology_list[idx1])
    topology_out_list.sort()

    return sorted(topology_list)


"""对每根信号线进行数据采集分析"""


# 指定最小pin number来筛选芯片(默认为4)
def get_sch_from_pin_number(SCH_brd_data, min_pin_number=4):
    min_pin_number = int(min_pin_number)

    SCH_content = SCH_brd_data.GetData()

    SCH_dict = dict()
    for line in SCH_content:
        if line[0] != '':
            # 键出现一次值为1，键出现两次值为2
            if SCH_dict.get(line[0]):
                SCH_dict[line[0]] += 1
            else:
                SCH_dict[line[0]] = 1

    sch_list = []
    for x in SCH_dict.keys():
        if SCH_dict[x] >= min_pin_number:
            sch_list.append(x)
    # 排序
    sch_list.sort()

    # print(sch_list)
    return sch_list


def data_analyze(excel_path):
    SCH_brd_data, Net_brd_data, diff_pair_brd_data, stackup_brd_data = report_filtrate()
    All_Net_List = getallnetlist(SCH_brd_data)
    diff_pair_dict, diff_list, single_ended_list, non_signal_net_list = net_separate(diff_pair_brd_data, All_Net_List)

    # 生成要check的起始芯片组
    start_sch_name_list = get_sch_from_pin_number(SCH_brd_data)

    SCH_object_list = SCH_detect(SCH_brd_data, non_signal_net_list)
    net_object_list = net_detect(Net_brd_data, SCH_object_list, non_signal_net_list)

    topology_dict = dict()
    ok_check_net_list, fail_check_net_list = list(), list()
    check_sch_ok_net_dict = dict()
    check_sch_fail_net_dict = dict()

    # ok_count = 0
    # 从起始芯片开始对每根线进行信息提取
    for check_sch_name in start_sch_name_list:
        check_net_list = get_connected_net_list_by_SCH_name(check_sch_name, SCH_object_list, net_object_list,
                                                            non_signal_net_list)
        # print(check_net_list)
        for check_net_name in check_net_list:
            try:
                # 遍历每个器件连接的所有线名
                # check_sch_name是符合最小pin脚个数的芯片，check_net_name是与芯片相接的信号线列表的遍历
                # if check_net_name == 'MXM_DPB_AUX_DN_C':
                #     print(check_net_name, check_sch_name)
                topology_list = topology_extract2(check_net_name, check_sch_name, SCH_object_list, net_object_list,
                                                  non_signal_net_list, excel_path)
                # print(11111, topology_list[0])
                # if check_net_name == 'DPC_AUX_DP_C':
                # print(11111, topology_list)
                topology_list = topology_list_format_simplified(topology_list[0], non_signal_net_list, All_Net_List)
                # print(2222, topology_1_list)
                # if check_net_name == 'DPC_AUX_DP_C':
                # print(22222, topology_1_list)
                topology_dict[check_net_name] = topology_list
                # print('topology_dict', topology_dict)
                if check_sch_ok_net_dict.get(check_sch_name) is None:
                    check_sch_ok_net_dict[check_sch_name] = [check_net_name]
                else:
                    check_sch_ok_net_dict[check_sch_name] += [check_net_name]
                ok_check_net_list.append(check_net_name)
            except:
                if check_sch_fail_net_dict.get(check_sch_name) is None:
                    check_sch_fail_net_dict[check_sch_name] = [check_net_name]
                else:
                    check_sch_fail_net_dict[check_sch_name] += [check_net_name]
                fail_check_net_list.append(check_net_name)

    # print(len(ok_check_net_list))
    # print(len(fail_check_net_list))
    # print(topology_dict)
    return topology_dict, ok_check_net_list, fail_check_net_list, diff_pair_dict, diff_list, single_ended_list


"""得到客户需要check的信号线的数据"""


def user_needed_signal_data():

    # 辨别出用户想匹配几类net_name
    parameter, param_path, output_path, excel_path = get_parameter()

    # 对parameter的格式进行容错处理
    if parameter:
        # 如果写成中文逗号
        if '，' in parameter:
            error_message = 'Please separate the parameters with a comma instead of Chinese commas'
            create_error_message(excel_path, error_message)
            raise FileNotFoundError

    # parameter为空
    else:
        error_message = 'No parameters filled in'
        create_error_message(excel_path, error_message)
        raise FileNotFoundError

    user_signal_org_list = parameter.strip().split(',')
    user_signal_org_list = [x.strip() for x in user_signal_org_list]

    topology_dict, _, _, diff_pair_dict, diff_list, single_ended_list = data_analyze(excel_path)
    sch_pin_list = list(topology_dict.keys())

    user_diff_out_list = []
    user_single_out_list = []

    for ind0 in range(len(user_signal_org_list)):
        user_signal_list = []
        user_no_repeat_signal_list = []
        user_diff_pair_signal_list = []
        user_single_signal_list = []

        middle_word_flag = True
        if user_signal_org_list[ind0].find('*') > -1:
            for ind1 in range(len(sch_pin_list)):
                if str(sch_pin_list[ind1]).find(str(user_signal_org_list[ind0][1:])) > -1 and not str(
                        sch_pin_list[ind1]).startswith(str(user_signal_org_list[ind0][1:])):
                    middle_word_flag = False
                    user_signal_list.append(sch_pin_list[ind1])

            # 如果没有word在中间或结尾的线
            if middle_word_flag:
                error_message = 'No signal line with {} in its name(not at the beginning)'.format(
                    str(user_signal_org_list[ind0][1:]))
                create_error_message(excel_path, error_message)
                raise FileNotFoundError

        else:
            #
            # print(22, user_signal_org_list[ind0])
            begin_flag = True
            for ind1 in range(len(sch_pin_list)):
                # print(sch_pin_list[ind][1])
                if str(sch_pin_list[ind1]).startswith(str(user_signal_org_list[ind0])):
                    # print(sch_pin_list[ind][1])
                    begin_flag = False
                    user_signal_list.append(sch_pin_list[ind1])

            # 如果没有以word开头的信号线
            if begin_flag:
                error_message = 'No signal line starting with {}'.format(str(user_signal_org_list[ind0]))
                create_error_message(excel_path, error_message)
                raise FileNotFoundError

        user_signal_list = sorted(user_signal_list)

        for x in user_signal_list:
            if x not in user_no_repeat_signal_list:
                user_no_repeat_signal_list.append(x)

        for value in diff_list:
            if value[0] in user_no_repeat_signal_list:
                user_diff_pair_signal_list.append(value)

        for x in user_no_repeat_signal_list:
            if x in single_ended_list:
                user_single_signal_list.append(x)

        user_diff_pair_signal_list = list(_flatten(user_diff_pair_signal_list))

        # 将信号线按照名称长度从小到大排列
        user_single_signal_list.sort(key=lambda x: len(x))

        user_diff_out_list.append(user_diff_pair_signal_list)
        user_single_out_list.append(user_single_signal_list)

    # print(111, user_diff_out_list)
    # print(222, user_single_out_list)

    return user_diff_out_list, user_single_out_list, topology_dict, diff_pair_dict, user_signal_org_list, \
           param_path, output_path, excel_path


"""生成供用户下载的报告"""


def create_excel():
    diff_net_name_list, single_net_name_list, topology_dict, diff_pair_dict, signal_list, \
    param_path, output_path, excel_path = user_needed_signal_data()
    # print(topology_dict)

    # 创建excel
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()
    next_table_position = 0

    for ind_net in range(len(diff_net_name_list)):
        diff_total_signal_list, diff_total_length_list, diff_total_via_list, diff_net_judge_total_signal = \
            get_total_signal(diff_net_name_list[ind_net], topology_dict)
        single_total_signal_list, single_total_length_list, single_total_via_list, single_net_judge_total_signal = \
            get_total_signal(single_net_name_list[ind_net], topology_dict)

        # 获取差分net的total mismatch
        diff_mismatch_list = total_mismatch(diff_total_signal_list, diff_total_length_list)
        # print(diff_total_signal_list)

        # 将要show的三个信息合并在一起
        diff_name_length_mismatch_list = zip(diff_total_signal_list, diff_total_length_list, diff_mismatch_list,
                                             diff_total_via_list)
        single_name_length_mismatch_list = zip(single_total_signal_list, single_total_length_list,
                                               single_total_via_list)

        diff_name_length_mismatch_dict = dict(zip(diff_net_judge_total_signal, diff_name_length_mismatch_list))
        single_name_length_mismatch_dict = dict(zip(single_net_judge_total_signal, single_name_length_mismatch_list))

        # print(len(diff_net_judge_total_signal))
        # print(len(diff_name_length_mismatch_list))
        # print(111, diff_net_judge_total_signal)
        # print(222, diff_name_length_mismatch_dict)
        diff_name_length_mismatch_list = get_one_way_signal(diff_net_judge_total_signal, diff_name_length_mismatch_dict,
                                                            diff_pair_dict)
        single_name_length_mismatch_list = get_one_way_signal(single_net_judge_total_signal,
                                                              single_name_length_mismatch_dict)

        # print(diff_name_length_mismatch_list)

        if ind_net > 0:
            next_table_position += max(len(prev_diff_net_list), len(prev_single_net_list)) + 3

        prev_diff_net_list = diff_name_length_mismatch_list

        prev_single_net_list = single_name_length_mismatch_list
        # title写入excel
        wb.sheets[0].range((1 + next_table_position, 1)).value = signal_list[ind_net]
        wb.sheets[0].range((1 + next_table_position, 1)).api.Font.Size = 22
        if diff_name_length_mismatch_list:

            # 创建表格
            wb.sheets[0].range((2 + next_table_position, 2)).value = '  Net Name(diff)'
            wb.sheets[0].range((2 + next_table_position, 3)).value = '   Total Length(mils)'
            wb.sheets[0].range((2 + next_table_position, 4)).value = '   Total Mismatch(mils)  '
            wb.sheets[0].range((2 + next_table_position, 5)).value = '   Via Number  '

            # 数据写入excel
            wb.sheets[0].range((3 + next_table_position, 2)).value = diff_name_length_mismatch_list

            # 标题UI
            wb.sheets[0].range((2 + next_table_position, 2), (2 + next_table_position, 5)).color = (45, 140, 168)
            wb.sheets[0].range((2 + next_table_position, 2),
                                      (2 + next_table_position, 5)).api.Font.Color = 0xFFFFFF
            wb.sheets[0].range((2 + next_table_position, 2)).row_height = 25
            wb.sheets[0].range((2 + next_table_position, 2), (2 + next_table_position, 5)).api.Font.Size = 16
            wb.sheets[0].range((2 + next_table_position, 3), (2 + next_table_position, 5)) \
                .api.HorizontalAlignment = -4152

            # 内容数据UI
            wb.sheets[0].range((3 + next_table_position, 2),
                                      (
                                          3 + next_table_position + len(diff_name_length_mismatch_list),
                                          5)).row_height = 18
            for x in range(int(len(diff_name_length_mismatch_list) / 2)):
                wb.sheets[0].range((3 + 2 * x + next_table_position, 2), (3 + 2 * x + next_table_position, 5)) \
                    .color = (218, 238, 243)
                wb.sheets[0].range((2 + 2 * x + next_table_position, 2), (2 + 2 * x + next_table_position, 5)) \
                    .api.VerticalAlignment = -4108

        if diff_name_length_mismatch_list and single_name_length_mismatch_list:

            # 创建表格
            wb.sheets[0].range((2 + next_table_position, 7)).value = '  Net Name(single)'
            wb.sheets[0].range((2 + next_table_position, 8)).value = '   Total Length(mils)  '
            wb.sheets[0].range((2 + next_table_position, 9)).value = '   Via Number  '

            # 数据写入excel
            wb.sheets[0].range((3 + next_table_position, 7)).value = single_name_length_mismatch_list

            # 标题UI
            wb.sheets[0].range((2 + next_table_position, 7), (2 + next_table_position, 9)).color = (45, 140, 168)
            wb.sheets[0].range((2 + next_table_position, 7), (2 + next_table_position, 9)) \
                .api.Font.Color = 0xFFFFFF
            wb.sheets[0].range((2 + next_table_position, 7)).row_height = 25
            wb.sheets[0].range((2 + next_table_position, 7), (2 + next_table_position, 9)).api.Font.Size = 16
            wb.sheets[0].range((2 + next_table_position, 8), (2 + next_table_position, 9)) \
                .api.HorizontalAlignment = -4152

            # 内容数据UI
            wb.sheets[0].range((3 + next_table_position, 7),
                                      (3 + next_table_position +
                                       len(single_name_length_mismatch_list), 9)).row_height = 18
            for x in range(int(math.ceil(len(single_name_length_mismatch_list) / 2.0))):
                wb.sheets[0].range((3 + 2 * x + next_table_position, 7), (3 + 2 * x + next_table_position, 9)) \
                    .color = (218, 238, 243)
                wb.sheets[0].range((2 + 2 * x + next_table_position, 7), (2 + 2 * x + next_table_position, 9)) \
                    .api.VerticalAlignment = -4108

        if diff_name_length_mismatch_list == [] and single_name_length_mismatch_list:

            # 创建表格
            wb.sheets[0].range((2 + next_table_position, 2)).value = '  Net Name(single)'
            wb.sheets[0].range((2 + next_table_position, 3)).value = '    Total Length(mils)  '
            wb.sheets[0].range((2 + next_table_position, 4)).value = '    Via Number  '

            # 数据写入excel
            wb.sheets[0].range((3 + next_table_position, 2)).value = single_name_length_mismatch_list

            # 标题UI
            wb.sheets[0].range((2 + next_table_position, 2), (2 + next_table_position, 4)).color = (45, 140, 168)
            wb.sheets[0].range((2 + next_table_position, 2),
                                      (2 + next_table_position, 4)).api.Font.Color = 0xFFFFFF
            wb.sheets[0].range((2 + next_table_position, 2)).row_height = 25
            wb.sheets[0].range((2 + next_table_position, 2), (2 + next_table_position, 4)).api.Font.Size = 16
            wb.sheets[0].range((2 + next_table_position, 3), (2 + next_table_position, 4)) \
                .api.HorizontalAlignment = -4152

            # 内容数据UI
            wb.sheets[0].range((3 + next_table_position, 2),
                                      (3 + next_table_position + len(single_name_length_mismatch_list), 4)) \
                .row_height = 18
            for x in range(int(math.ceil(len(single_name_length_mismatch_list) / 2.0))):
                wb.sheets[0].range((3 + 2 * x + next_table_position, 2), (3 + 2 * x + next_table_position, 4)) \
                    .color = (218, 238, 243)
                wb.sheets[0].range((2 + 2 * x + next_table_position, 2), (2 + 2 * x + next_table_position, 4)) \
                    .api.VerticalAlignment = -4108

    wb.sheets[0].autofit('c')

    wb.save(excel_path)
    wb.close()
    app.quit()


if __name__ == '__main__':
    create_excel()
