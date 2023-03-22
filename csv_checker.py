import csv



def insert_error_info_to_dict(_err_dict, line_number, error_key, error_description):
    if len(_err_dict.items()) == 0 or _err_dict is None:
        _err_dict = {str(line_number): {error_key: [error_description]}}
    elif str(line_number) in _err_dict.keys():
        if error_key in _err_dict[str(line_number)].keys():
            _err_dict[str(line_number)][error_key].append(error_description)
        else:
            _err_dict[str(line_number)][error_key] = [error_description]
    else:
        _err_dict[str(line_number)] = {error_key: [error_description]}
        pass

    return _err_dict


def check_integrity_of_order_file(_file_path, _encoding='utf-8'):
    is_file_valid = False
    error_dictionary = {}

    #print("csv_checker: Path - " + _file_path)


    f = open(_file_path, 'r', encoding=_encoding)
    lines = f.readlines()
    
    for line_num, line in enumerate(lines):
        IS_LINE_READY_TO_CHECK_VALUE = True


        if line_num == 0:
            # IS HEADER
            header = line.rstrip().split(',')
            header_count = len(header)
            #print("")
            #print("========== HEADER ==========")
            #print(header)
            #print("")
            continue

        tmp_line = line.rstrip()

        #curr_line_comma_cnt = tmp_line.count(",")
        curr_line_element_cnt = tmp_line.count(",") + 1

        if curr_line_element_cnt < header_count:
            #print("ERROR: Number of element is not sufficient at line {0}".format(line_num + 1))
            error_dictionary = insert_error_info_to_dict(error_dictionary, line_num+1, "ELEM_NOT_SUFF", "해당 줄에 정보 부족")
            IS_LINE_READY_TO_CHECK_VALUE = False
        elif curr_line_element_cnt > header_count:
            #print("WARNING: Need to delete some comma at line {0}".format(line_num + 1))
            pass

        single_quotation_mark = "'"
        double_quotation_mark = '"'
        
        single_quotation_indices = [pos for pos, char in enumerate(line) if char == single_quotation_mark]
        double_quotation_indices = [pos for pos, char in enumerate(line) if char == double_quotation_mark]

        if len(single_quotation_indices) % 2 == 1:
            #print("ERROR: Single quotation not matches at line {0}".format(line_num + 1))
            error_dictionary = insert_error_info_to_dict(error_dictionary, line_num+1, "S_QUOT_NOT_MATCH", "홑따옴표 쌍이 맞지 않음 (처리 전)")
            IS_LINE_READY_TO_CHECK_VALUE = False
        if len(double_quotation_indices) % 2 == 1:
            #print("ERROR: Double quotation not matches at line {0}".format(line_num + 1))
            error_dictionary = insert_error_info_to_dict(error_dictionary, line_num+1, "D_QUOT_NOT_MATCH", "쌍따옴표 쌍이 맞지 않음 (처리 전)")
            IS_LINE_READY_TO_CHECK_VALUE = False

        part_string = tmp_line
        # Remove single quotation mark
        while len(single_quotation_indices) > 0 and IS_LINE_READY_TO_CHECK_VALUE:
            start_index = single_quotation_indices[0]
            end_index = single_quotation_indices[1]

            part_string = [part_string[:start_index], part_string[start_index:end_index+1], part_string[end_index+1:]]

            front_string = part_string[0]
            middle_string = part_string[1]
            end_string = part_string[2]

            middle_string = middle_string.replace(',', ' ')
            middle_string = middle_string[1:-1]

            part_string = front_string+middle_string+end_string

            single_quotation_indices = [pos for pos, char in enumerate(part_string) if char == single_quotation_mark]

            if len(single_quotation_indices) == 1:
                #print("ERROR: Single quotation not matches at line {0}".format(line_num + 1))
                error_dictionary = insert_error_info_to_dict(error_dictionary, line_num+1, "S_QUOT_NOT_MATCH", "홑따옴표 쌍이 맞지 않음 (처리 후)")
                IS_LINE_READY_TO_CHECK_VALUE = False


        # Remove double quotation mark
        while len(double_quotation_indices) > 0 and IS_LINE_READY_TO_CHECK_VALUE:
            start_index = double_quotation_indices[0]
            end_index = double_quotation_indices[1]

            part_string = [part_string[:start_index], part_string[start_index:end_index+1], part_string[end_index+1:]]

            front_string = part_string[0]
            middle_string = part_string[1]
            end_string = part_string[2]

            middle_string = middle_string.replace(',', ' ')
            middle_string = middle_string[1:-1]

            part_string = front_string+middle_string+end_string

            double_quotation_indices = [pos for pos, char in enumerate(part_string) if char == double_quotation_mark]

            if len(double_quotation_indices) == 1:
                #print("ERROR: Double quotation not matches at line {0}".format(line_num + 1))
                error_dictionary = insert_error_info_to_dict(error_dictionary, line_num+1, "D_QUOT_NOT_MATCH", "쌍따옴표 쌍이 맞지 않음 (처리 후)")
                IS_LINE_READY_TO_CHECK_VALUE = False
        
        curr_line_element_cnt = part_string.count(",") + 1


        if curr_line_element_cnt != header_count:
            #print("ERROR: Excess comma at line {0}".format(line_num + 1))
            error_dictionary = insert_error_info_to_dict(error_dictionary, line_num+1, "EXCESS_COMMA", "쉼표가 너무 많음")
            IS_LINE_READY_TO_CHECK_VALUE = False

        '''        
        if IS_LINE_READY_TO_CHECK_VALUE:
            #print("L{0}: READY TO CHECK VALUE".format(line_num+1))

            # Check is number valid


            # Cehck 
        '''


        
    f.close()

    print(error_dictionary)
    if len(error_dictionary.keys()) == 0:
        is_file_valid = True
    else:
        is_file_valid = False

    return is_file_valid, error_dictionary













if __name__ == "__main__":
    kt_file_path = './order_data/20210311/raw/1031_01_20210311_KT.CSV'
    mns_file_path = './order_data/20210311/raw/1031_01_20210311_MnS.CSV'
    
    
    #kt_file_path = './order_data/20210309/raw/1031_08_20210309_KT.csv'
    #mns_file_path = './order_data/20210309/raw/1031_08_20210309_MnS.csv'
    
    #file_path = './order_data/20210311/raw/1031_01_20210311_KT.CSV'
    #file_path = './order_data/20210311/raw/1031_01_20210311_KT.CSV'

    check_integrity_of_order_file(kt_file_path)
    check_integrity_of_order_file(mns_file_path)