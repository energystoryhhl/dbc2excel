import re
import xlwt
import wx

length_of_BO = 5
str_of_BO = 'BO_'
str_of_SG = 'SG_'
location_of_bo_type = 0
location_of_bo_id = 1
location_of_bo_message_name = 2
location_of_bo_dlc = 3
location_of_bo_transmitter = 4

location_of_sg_type = 0
location_of_sg_name = 1
location_of_sg_s_bit_size = 3
location_of_sg_factor = 4
location_of_sg_max_min = 5
location_of_sg_unit = 6
location_of_sg_receiver = 7

if_show_global = 0

excel_page_name = "Matrix"
tittle_row = 0
signal_name_col = 6

excel_tittle = ('Msg Name\n报文名称', 'Msg Type\n报文类型', 'Msg ID\n报文标识符', 'Msg Send Type\n报文发送类型',
                'Msg Cycle Time (ms)\n报文周期时间', 'Msg Length (Byte)\n报文长度', 'Signal Name\n信号名称',
                'Signal Description\n信号描述', "Byte Order\n排列格式(Intel/Motorola)", "Start Byte\n起始字节",
                "Start Bit\n起始位", "Signal Send Type\n信号发送类型", "Bit Length (Bit)\n信号长度", "Date Type\n数据类型",
                "Resolution\n精度", "Offset\n偏移量", "Signal Min. Value (phys)\n物理最小值", "Signal Max. Value (phys)\n物理最大值",
                "Signal Min. Value (Hex)\n总线最小值", "Signal Max. Value (Hex)\n总线最大值", "Initial Value (Hex)\n初始值",
                "Invalid Value(Hex)\n无效值", "Inactive Value (Hex)\n非使能值", "Unit\n单位", "Signal Value Description\n信号值描述",
                "Msg Cycle Time Fast(ms)\n报文发送的快速周期(ms)", "Msg Nr. Of Reption\n报文快速发送的次数", "Msg Delay Time(ms)\n报文延时时间(ms)",
                )


def set_style( color = 0, bold = False,italic = False):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    # 字体类型：比如宋体、仿宋也可以是汉仪瘦金书繁
    font.name = 'Arial'
    # 是否为粗体
    font.bold = bold
    # 设置字体颜色
    font.colour_index = 0
    # 字体大小
    font.height = 200
    # 字体是否斜体
    font.italic = italic
    # 字体下划,当值为11时。填充颜色就是蓝色
    font.underline = 0
    # 字体中是否有横线struck_out
    font.struck_out =False
    # 定义格式
    style.font = font

    ##
    borders = xlwt.Borders()
    borders.left = 0
    borders.right = 0
    borders.top = 0
    borders.bottom = 0
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.borders = borders
    ##

    alignment = xlwt.Alignment()  # 创建居中
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 可取值: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 可取值: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行
    style.alignment = alignment  # 给样式添加文字居中属性

    # 设置背景颜色
    pattern = xlwt.Pattern()
    # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # 背景颜色
    pattern.pattern_fore_colour = color
    style.pattern = pattern

    return style


#DBC LOAD CLASS
class DbcLoad(object):
    def __init__(self, dbc_name_in):
        self.dbc_fd = open(dbc_name_in, 'r')
        if self.dbc_fd.readable():
            if(if_show_global):
                print('>>>DBC File loaded!')
            self.num_of_bo = 0
            self.num_of_sg = 0
            self.dbc_list = []
            self.dbc_name = dbc_name_in
            self.dbc_cycle_time = {}
            self.tran_recv = set()
            self.tran_recv_list = []
            ##excel生成条件变量
            self.if_sig_desc = True
            self.if_sig_val_desc = True
            self.val_description_max_number = 70
            self.if_start_val = True
            self.if_recv_send = True
        else:
            if (if_show_global):
                print("DEC file load failed!")

    def cm_put(self, canid, sg_name, comment):
        dbc_length = len(self.dbc_list)
        i = 0
        j = 0
        while i < dbc_length:
            flag = 0
            for bo_line in self.dbc_list[i]:
                if 'message_id' in bo_line:
                    if bo_line['message_id'] == canid and flag == 0:
                    #bo_index_len = len(self.dbc_list[i])
                        flag = 1
                if flag == 1:
                    if 'signal_name' in bo_line:
                        if bo_line['signal_name'] == sg_name:
                            bo_line['comment'] = comment
                            break
            i = i+1
    def put_inedx(self, canid, sg_name, index, comment):
        dbc_length = len(self.dbc_list)
        i = 0
        j = 0
        while i < dbc_length:
            flag = 0
            for bo_line in self.dbc_list[i]:
                if 'message_id' in bo_line:
                    if bo_line['message_id'] == canid and flag == 0:
                    #bo_index_len = len(self.dbc_list[i])
                        flag = 1
                if flag == 1:
                    if 'signal_name' in bo_line:
                        if bo_line['signal_name'] == sg_name:
                            bo_line[index] = comment
                            break
            i = i+1

    def bit_mask(self, nbits):
        ret = 0
        for i in range(nbits ):
            ret = ret * 2 + 1
        return ret


    def parse_dbc(self, if_show, if_sig_desc,if_sig_val_desc,val_description_max_number):
        self.if_sig_desc = if_sig_desc
        self.if_sig_val_desc = if_sig_val_desc
        self.val_description_max_number = val_description_max_number
        line_list = self.dbc_fd.readlines()
        #读取cycletime
        i = 0
        length = len(line_list)
        while i < length:
            if len(line_list[i].split()) > 2:
                if line_list[i].split()[0] == 'BA_' and line_list[i].split()[1] == '"GenMsgCycleTime"':
                    #查找到了相应的cycletime的line
                    self.dbc_cycle_time[line_list[i].split()[3]] = int(re.sub(';', '', line_list[i].split()[4]))
                    if (if_show_global):
                        print(line_list[i].split())
            i += 1

        #读取具体讯息
        i = 0
        length = len(line_list)
        while i < length:
            if len(line_list[i].split()) == length_of_BO:
                if line_list[i].split()[0] == str_of_BO: #检测BOlist
                    bo_list = []
                    bo_line = line_list[i].split()
                    if if_show:
                        print(bo_line)
                    ##对BO_LIST的操作
                    self.num_of_bo += 1
                    ##构造bo字典
                    bo_dict = {}
                    bo_dict['type'] = bo_line[location_of_bo_type]
                    bo_dict['message_id'] = int(bo_line[location_of_bo_id])
                    bo_dict['message_name'] = re.sub(':', '', bo_line[location_of_bo_message_name])
                    bo_dict['message_size'] = int(bo_line[location_of_bo_dlc])
                    bo_dict['transmitter'] = bo_line[location_of_bo_transmitter]
                    if str(bo_dict['message_id']) in self.dbc_cycle_time:
                        bo_dict['cycle_time'] = self.dbc_cycle_time[str(bo_dict['message_id'])]
                    bo_list.append(bo_dict)                     #加入bo_list中
                    if if_show:
                        print('转换得到：'+str(bo_dict))
                    ##对BO_LIST的操作
                    if line_list[i+1] != "\n":
                        i += 1
                        while line_list[i].split()[0] == str_of_SG : #循环读取SGlist
                            sg_list = line_list[i].split()
                            if if_show:
                                print(sg_list)
                            ##对SG_LIST的操作
                            self.num_of_sg += 1
                            sg_dict = {}
                            sg_dict['type'] = sg_list[location_of_sg_type]
                            sg_dict['signal_name'] = sg_list[location_of_sg_name]
                            if(sg_list[2] != ":"):
                                break
                            end_of_start = sg_list[location_of_sg_s_bit_size].find('|')
                            end_of_size = sg_list[location_of_sg_s_bit_size].find('@')
                            sg_dict['start_bit'] = int(sg_list[location_of_sg_s_bit_size][0:end_of_start])
                            sg_dict['signal_size'] = int(sg_list[location_of_sg_s_bit_size][end_of_start+1:end_of_size])
                            sg_dict['byte_order'] = int(sg_list[location_of_sg_s_bit_size][end_of_size+1])
                            sg_dict['value_type'] = sg_list[location_of_sg_s_bit_size][end_of_size+2]
                            if sg_dict['value_type'] == '+':
                                sg_dict['value_type'] = 0
                            else:
                                sg_dict['value_type'] = 1
                            end_of_factor = sg_list[location_of_sg_factor].find(',')
                            end_of_offset = sg_list[location_of_sg_factor].find(')')
                            sg_dict['factor'] = float(sg_list[location_of_sg_factor][1:end_of_factor])
                            sg_dict['offset'] = float(sg_list[location_of_sg_factor][end_of_factor+1:end_of_offset])
                            end_of_min = sg_list[location_of_sg_max_min].find('|')
                            end_of_max = sg_list[location_of_sg_max_min].find(']')
                            sg_dict['minimum'] = float(sg_list[location_of_sg_max_min][1:end_of_min])
                            sg_dict['maximum'] = float(sg_list[location_of_sg_max_min][end_of_min+1:end_of_max])
                            sg_dict['unit'] = re.sub('"', '', sg_list[location_of_sg_unit])
                            sg_dict['receiver'] = sg_list[location_of_sg_receiver]
                            bo_list.append(sg_dict)     #加入bo_list中
                            if if_show:
                                print('起始位：'+str(sg_dict['start_bit']), end=' ')
                                print('长度：' + str(sg_dict['signal_size']), end=' ')
                                print('格式：' + str(sg_dict['byte_order']), end=' ')
                                print('是否有符号：' + str(sg_dict['value_type']), end=' ')
                                print('factor：' + str(sg_dict['factor']), end=' ')
                                print('offset：' + str(sg_dict['offset']), end=' ')
                                print('最小值：' + str(sg_dict['minimum']), end=' ')
                                print('最大值：' + str(sg_dict['maximum']), end=' ')
                                print('单位：' + sg_dict['unit'], end=' ')
                                print('接收单元：' + sg_dict['receiver'], end=' ')
                                print()

                            ##对SG_LIST的操作
                            i += 1
                            if len(line_list[i].split()) == 0:      #不是SGlist 则break！
                                if if_show:
                                    print()
                                break
                            elif line_list[i].split()[0] != str_of_SG:
                                break
                    self.dbc_list.append(bo_list)
                    if if_show:
                        print("DBC数量:" + str(self.num_of_bo))
            i += 1

        #读取CM_ SG_Signal Description
        if(self.if_sig_desc):
            i = 0
            length = len(line_list)
            bo_id = 0
            sg_name = ''
            while i < length:
                if len(line_list[i].split()) > 2:
                    if line_list[i].split()[0] == 'CM_' and line_list[i].split()[1] == 'SG_':#查找到了相应的CM的line
                        bo_id = int(line_list[i].split()[2])
                        sg_name = line_list[i].split()[3]#获取canid和sgname
                        #截取cm
                        if line_list[i][-2] != ';':
                            comment = str(line_list[i].split('"')[-1])
                            t = 0
                            while(True):
                                t = t + 1
                                comment = comment + line_list[i+t]
                                if( line_list[i + t] != "\n"):
                                    if (line_list[i + t][-2] == ';'):
                                        break

                                # if line_list[i][-2] == ';':
                                #     break
                        else:
                            comment = line_list[i].split('"')[-2]
                            #print(line_list[i])
                        #self.cm_put(bo_id, sg_name, comment)
                        self.put_inedx(bo_id,sg_name,"comment",comment)
                        if if_show:
                            print(str(bo_id)+' '+sg_name)
                            print(comment)
                i += 1


        #读取START_VALUE
        if(self.if_start_val ):
            i = 0
            length = len(line_list)
            bo_id = 0
            sg_name = ''
            while i < length:
                if len(line_list[i].split()) > 2:
                    if line_list[i].split()[0] == 'BA_' and line_list[i].split()[1] == '"GenSigStartValue"':#查找到了相应的START_VALUE
                        bo_id = int(line_list[i].split()[3])
                        sg_name = line_list[i].split()[4]
                        inital_value = hex(int(float((line_list[i].split()[5].rstrip(';')))))
                        self.put_inedx(bo_id, sg_name, "inital_value", inital_value)
                        if if_show:
                            print(str(bo_id)+' '+sg_name+str(inital_value))
                            print(comment)
                i += 1

        #信号值描述
        if(self.if_sig_val_desc):
            i = 0
            length = len(line_list)
            bo_id = 0
            sg_name = ''
            while i < length:
                if len(line_list[i].split()) > 2:
                    if line_list[i].split()[0] == 'VAL_':#查找到了相应的描述
                        bo_id = int(line_list[i].split()[1])
                        sg_name = line_list[i].split()[2]
                        j = 1
                        val_des_list = []
                        #数字
                        val_des_list.append(str(hex(int(line_list[i].split()[3])))+': ')
                        while line_list[i].split()[j] != ' ;\n':
                            val_des_list.append(line_list[i].split('"')[j]+'\n')
                            j = j+1
                            if line_list[i].split('"')[j] == ' ;\n':
                                break
                            #数字
                            val_des_list.append(str(hex(int(line_list[i].split('"')[j])))+': ')
                            j = j+1
                        if len(val_des_list) <= self.val_description_max_number:
                            self.put_inedx(bo_id, sg_name, "val_description", val_des_list)
                        #val_des_list = []
                i += 1

        #读取发送者和接收者
        if(self.if_recv_send):
            len_of_dbc_list = len(self.dbc_list)
            i = 0
            while i < len_of_dbc_list:
                for bo_line in self.dbc_list[i]:
                    if 'transmitter' in bo_line:
                        if bo_line['transmitter'].find(',') != -1:
                            self.tran_recv.add(bo_line['transmitter'].split(','))
                        else:
                            self.tran_recv.add(bo_line['transmitter'])
                        #self.tran_recv.add(bo_line['transmitter'])
                    if 'receiver' in bo_line:
                        if bo_line['receiver'].find(',')!= -1:
                            self.tran_recv.add(bo_line['receiver'].split(',')[0])
                            self.tran_recv.add(bo_line['receiver'].split(',')[1])
                        else:
                            self.tran_recv.add(bo_line['receiver'])
                        #self.tran_recv.add(bo_line['receiver'])
                i += 1
            for each in self.tran_recv:
                self.tran_recv_list.append(each)
            if (if_show_global):
                print(self.tran_recv_list)

        ############打印结果
        if if_show:
            print(self.tran_recv_list)
        if if_show:
            print('dbc文件一共有{}个BO信号'.format(self.num_of_bo))
        return self.dbc_list

    def dbc_info(self):
        length = len(self.dbc_list)
        if (if_show_global):
            print("DBC中一共的信号：" + str(length))
        i = 0
        while i < length:
            for line in self.dbc_list[i]:
                print(line)
            print()
            i = i + 1

    def dbc_head_code_gen(self):
        end_of_dbc = str(self.dbc_name).find('.')
        head_code_fd = open(str(self.dbc_name[0:end_of_dbc])+'.h', 'w+')
        head_code_fd.write('#ifndef '+'__'+str(self.dbc_name[0:end_of_dbc]).upper() + '_'+'H__'+'\n')
        head_code_fd.write('#define ' + '__' + str(self.dbc_name[0:end_of_dbc]).upper() + '_' + 'H__'+'\n')

        #编译所有BO进行结构体定义
        bo_list_struct = []
        length = len(self.dbc_list)
        if (if_show_global):
            print("DBC中一共的信号：" + str(length))
        i = 0
        while i < length:
            # 定义struct结构体
            for line in self.dbc_list[i]: #line为每个bo中的字典
                if 'message_id' in line:
                    head_code_fd.write('struct ' + 'BO_' + str(line['message_id']) + '{' + '\n')
                    head_code_fd.write('    '+'double cycle_time'+';'+'\n')
                    bo_list_struct.append('struct ' + 'BO_' + str(line['message_id']) + ' message_' + str(line['message_id']) )
            #定义struct内部的成员
            for line_sg in self.dbc_list[i]: #line为每个bo中的字典
                if 'message_id' in line_sg:
                    pass
                else:
                    head_code_fd.write('    '+ 'double ' + line_sg['signal_name'] + ';'+'\n')
            # 定义struct结构体尾部
            head_code_fd.write( '};' + '\n\n')
            i += 1
        #生成一个总的bolist结构
        head_code_fd.write(' struct ' + 'BO_List' +  '{' + '\n')
        for i in bo_list_struct:
            head_code_fd.write('    '+i + ';' + '\n')
        head_code_fd.write('};' + '\n')

        #接收buffer结构体的定义
        bo_list_struct = []
        length = len(self.dbc_list)
        if (if_show_global):
            print("DBC中一共的信号：" + str(length))
        i = 0
        while i < length:
            # 定义struct结构体
            for line in self.dbc_list[i]: #line为每个bo中的字典
                if 'message_id' in line:
                    head_code_fd.write('struct ' + 'BO_' + str(line['message_id']) +'_recv'+ '{' + '\n')
                    bo_list_struct.append('struct ' + 'BO_' + str(line['message_id']) + ' message_' + str(line['message_id'])+'_recv' )
            #定义struct内部的成员
            for line_sg in self.dbc_list[i]: #line为每个bo中的字典
                if 'message_id' in line_sg:
                    pass
                else:
                    head_code_fd.write('    '+ 'double ' + line_sg['signal_name'] + ';'+'\n')
            # 定义struct结构体尾部
            head_code_fd.write( '};' + '\n\n')
            i += 1
        #生成一个总的bolist结构
        head_code_fd.write(' struct ' + 'BO_List_Recv' +  '{' + '\n')
        for i in bo_list_struct:
            head_code_fd.write('    '+i + ';' + '\n')
        head_code_fd.write('};' + '\n')

        #endif结尾
        head_code_fd.write('#endif '+'\n')

    #解析代码生成
    def dbc_parse_code_gen(self):
        end_of_dbc = str(self.dbc_name).find('.')
        parse_code_fd = open('parse.c', 'w+')
        #包含头文件
        parse_code_fd.write('#include ' + '"'+str(self.dbc_name[0:end_of_dbc])+'.h' + '"' + '\n')
        #解析代码生成
        parse_code_fd.write('int parse(struct can_frame * frame, struct BO_List* bo_list) \n')
        parse_code_fd.write('{ \n')
        parse_code_fd.write('int message_id = frame->can_id \n')
        #switch case 头
        parse_code_fd.write('switch (message_id) {\n')
        #具体解析

        #遍历dbc_list
        length = len(self.dbc_list)
        i = 0
        while i < length:
            for bo_sg in self.dbc_list[i]: #bo_sg为每个bo里的sg和bo字典
                if 'message_name' in bo_sg:
                    #case 头
                    parse_code_fd.write('case ' + str(bo_sg['message_id']) + ':\n')
                    messageg_id = 'message_id_' + str(bo_sg['message_id'])
                    parse_code_fd.write('       double tmp = 0 \n')
                    #具体解析过程
                if 'signal_name' in bo_sg:
                    if bo_sg['signal_size'] <= 8:
                        parse_code_fd.write('       tmp = 0 \n')
                        parse_code_fd.write('       tmp = ')
                        location = int(bo_sg['start_bit'] / 8)
                        mask_bit = self.bit_mask(bo_sg['signal_size'])
                        move_size = (bo_sg['start_bit'] % 8) - bo_sg['signal_size'] + 1
                        #初步提取
                        parse_code_fd.write('(frame->' + 'data[' + str(location) + ']' + '&' + '(' + str(mask_bit) + '<<' + str(move_size) +'))>>'+str(move_size) + ';\n')
                        #是否有符号
                        if bo_sg['value_type']:
                            #8字节：
                            parse_code_fd.write('       tmp = (signed char)tmp; \n')

                        #factor and offset
                        parse_code_fd.write('       tmp = tmp *  ' + str(bo_sg['factor']) + '+'+ str(bo_sg['offset']) +  ';\n')

                        #Max and Min
                        parse_code_fd.write('       if(tmp > {}) \n'.format(str(bo_sg['maximum'])))
                        parse_code_fd.write('           tmp = {};\n'.format(str(bo_sg['maximum'])))
                        parse_code_fd.write('       if(tmp < {}) \n'.format(str(bo_sg['minimum'])))
                        parse_code_fd.write('           tmp = {};\n'.format(str(bo_sg['minimum'])))
                        parse_code_fd.write('       '+'bo_list->' + messageg_id + '->' + bo_sg['signal_name'] + ' = tmp ;\n')


                    #break 尾部
            parse_code_fd.write('       '+'break;\n')
            parse_code_fd.write('\n')
            i += 1
        #switch case 结束花括号
        parse_code_fd.write('   '+'default:return -1;\n')
        parse_code_fd.write('}\n')
        #parse函数结束花括号
        parse_code_fd.write('\n}\n')

    def dbc_define_gen(self):
        length = len(self.dbc_list)
        i = 0
        dbc_define_fd = open(self.dbc_name[0:str(self.dbc_name).find('.')] + '_define.h', 'w+')
        while i < length:
            j = -1
            for bo_sg in self.dbc_list[i]: #bo_sg为每个bo里的sg和bo字典
                if 'signal_name' in bo_sg :
                    dbc_define_fd.write('#define ' + str(bo_sg['signal_name']) +'   BO_List[{}]->SG_List[{}]->value'.format(str(i), str(j))   + '\n')
                    dbc_define_fd.write('#define ' + str(bo_sg['signal_name'])+'_s' + '   BO_List[{}]->SG_List[{}]->value_send'.format(str(i), str(j))   + '\n')
                j += 1
            dbc_define_fd.write('\n')

            i += 1

    def dbc_excel_gen(self):
        if (if_show_global):
            print(">>>")
            print(self.dbc_list)
        book = xlwt.Workbook(encoding='utf-8')
        sheet = book.add_sheet(excel_page_name, cell_overwrite_ok = True)
        row_counter = 0

        #write tittle
        tittle_len = len(excel_tittle)
        i = 0
        while i < tittle_len:
            sheet.write(tittle_row, i, excel_tittle[i], set_style(0x28, True, True))
            i = i+1
        tran_recv_len = len(self.tran_recv_list)
        tmp = i + tran_recv_len
        j = 0
        while i < tmp:
            sheet.write(tittle_row, i, self.tran_recv_list[j], set_style(0x28, True, True))
            i += 1
            j += 1

        #调整行距离
        col0 = sheet.col(0)
        col0.width = 700*20
        col0 = sheet.col(2)
        col0.width = 200*20
        col0 = sheet.col(6)
        col0.width = 400*20
        col0 = sheet.col(7)
        col0.width = 700*20
        col0 = sheet.col(24)
        col0.width = 700*20
        #write BO units
        dbc_length = len(self.dbc_list)
        style = set_style(255, True, True)
        style_index = set_style(255, False, False)
        # alignment = xlwt.Alignment()  # 创建居中
        # alignment.horz = xlwt.Alignment.HORZ_CENTER  # 可取值: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        # alignment.vert = xlwt.Alignment.VERT_CENTER  # 可取值: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        # alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行
        # style.alignment = alignment  # 给样式添加文字居中属性
        i = 0
        while i < dbc_length:#dbc_length
            for bo_unit in self.dbc_list[i]:#for bo_unit in self.dbc_list[i]:
                if 'message_id' in bo_unit:
                    row_counter = row_counter + 1
                    sheet.write(tittle_row + row_counter, 0, bo_unit['message_name'], set_style(0x0D, True, True))
                    sheet.write(tittle_row + row_counter, 1, 'normal', set_style(0x28, True, True))
                    sheet.write(tittle_row + row_counter, 2, str(hex(bo_unit['message_id'])), set_style(0x0D, True, True))
                    if 'cycle_time' in bo_unit:
                        sheet.write(tittle_row + row_counter, 3, 'cycle', set_style(0x28, True, True))
                        sheet.write(tittle_row + row_counter, 4, bo_unit['cycle_time'], set_style(0x0D, True, True))
                    else:
                        sheet.write(tittle_row + row_counter, 3, 'NULL', set_style(0x28, True, True))
                        sheet.write(tittle_row + row_counter, 4, 'NULL', set_style(0x0D, True, True))
                    sheet.write(tittle_row + row_counter, 5, bo_unit['message_size'], set_style(0x28, True, True))
                    #发送者和接收者
                    k = 0
                    tran_recv_len = len(self.tran_recv_list)
                    while k < tran_recv_len:
                        if 'transmitter' in bo_unit:
                            if bo_unit['transmitter'].find(self.tran_recv_list[k]) != -1:
                                sheet.write(tittle_row + row_counter, signal_name_col + 22 + k, 'S', style_index)
                            else:
                                sheet.write(tittle_row + row_counter, signal_name_col + 22 + k, '', style_index)
                        k += 1
                if 'signal_name' in bo_unit:
                    row_counter = row_counter + 1
                    #信号名称
                    sheet.write(tittle_row + row_counter, signal_name_col, bo_unit['signal_name'],style)
                    #信号描述
                    if 'comment' in bo_unit:
                        sheet.write(tittle_row + row_counter, signal_name_col + 1, bo_unit['comment'],style_index)
                    else:
                        sheet.write(tittle_row + row_counter, signal_name_col + 1, '', style_index)
                    #字节序
                    if bo_unit['byte_order'] == 0:
                        sheet.write(tittle_row + row_counter, signal_name_col + 2, "Motorola LSB",style_index)
                    else:
                        sheet.write(tittle_row + row_counter, signal_name_col + 2, "Intel",style_index)
                    #起始字节
                    sheet.write(tittle_row + row_counter, signal_name_col + 3, str(int(bo_unit['start_bit'] / 8)),style_index)
                    #起始位
                    sheet.write(tittle_row + row_counter, signal_name_col + 4, str(bo_unit['start_bit']),style_index)
                    #循环类型
                    sheet.write(tittle_row + row_counter, signal_name_col + 5, 'cycle',style_index)
                    #信号类型
                    sheet.write(tittle_row + row_counter, signal_name_col + 6, str(bo_unit['signal_size']),style_index)
                    #数据类型
                    if bo_unit['value_type'] == 0:
                        sheet.write(tittle_row + row_counter, signal_name_col + 7, 'unsigned',style_index)
                    else:
                        sheet.write(tittle_row + row_counter, signal_name_col + 7, 'signed',style_index)
                    #信号精度
                    sheet.write(tittle_row + row_counter, signal_name_col + 8, str(bo_unit['factor']),style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 9, str(bo_unit['offset']),style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 10, str(bo_unit['minimum']),style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 11, str(bo_unit['maximum']),style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 12, '0x00',style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 13, str(hex(2**bo_unit['signal_size'] - 1)),style_index)
                    #初始值
                    if 'inital_value' in bo_unit:
                        sheet.write(tittle_row + row_counter, signal_name_col + 14, bo_unit['inital_value'],style_index)
                    else:
                        sheet.write(tittle_row + row_counter, signal_name_col + 14, '',style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 15, '',style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 16, '',style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 17, bo_unit['unit'],style_index)
                    #信号描述
                    if 'val_description' in bo_unit:
                        sheet.write(tittle_row + row_counter, signal_name_col + 18, bo_unit['val_description'],style_index)
                    else:
                        sheet.write(tittle_row + row_counter, signal_name_col + 18, '',
                                    style_index)
                    #周期次数延时
                    sheet.write(tittle_row + row_counter, signal_name_col + 19, '',style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 20, '',style_index)
                    sheet.write(tittle_row + row_counter, signal_name_col + 21, '',style_index)
                    #发送者和接收者
                    k = 0
                    tran_recv_len = len(self.tran_recv_list)
                    while k < tran_recv_len:
                        if 'receiver' in bo_unit:
                            if bo_unit['receiver'].find(self.tran_recv_list[k]) != -1:
                                sheet.write(tittle_row + row_counter, signal_name_col + 22 + k, 'R', style_index)
                            else:
                                sheet.write(tittle_row + row_counter, signal_name_col + 22 + k, '', style_index)
                        k += 1
                        #book.save(self.dbc_name.replace('.', '_') + '.xls')
            i = i+1

        book.save(self.dbc_name.replace('.', '_') + '.xls')

    def dbc2excel(self, filepath,if_sig_desc,if_sig_val_desc,val_description_max_number):
        self.dbc_fd = open(filepath, 'r')
        self.dbc_name = filepath.split("\\")[-1]
        self.parse_dbc(0,if_sig_desc,if_sig_val_desc,val_description_max_number)
        self.dbc_excel_gen()





if __name__ == "__main__":
    dbc_cls = DbcLoad('Tx2TryDBC.dbc')
    dbc_cls.dbc2excel('Tx2TryDBC.dbc')
    # dbc_cls.parse_dbc(1)
    # dbc_cls.dbc_info()
    # dbc_cls.dbc_excel_gen()
    ###############################
    #dbc_cls.dbc_head_code_gen()
    #dbc_cls.dbc_parse_code_gen()
    #dbc_cls.dbc_define_gen()




















