import sys
import math
import shutil
import pandas as pd

from openpyxl import load_workbook

import file_concat


EXCEL_HEADER = ['물류센터명','운송타입','배송지명','배송지 주소','배송지 상세주소',
                '주문유형','위도(Y)','경도(X)','운송부피(CBM)','운송중량(kg)',
                '박스수량','상품분류','상품명','상품수량','개별중량(kg)',
                '개별원가','자체관리코드','하차유형','점착시간','점착시작시간',
                '점착종료시간','회피시간','배송차량','배송순서']
ERP_HEADER = ['CENTER_NM','ORDER_TYPE','LOCATION_NM','ADDRESS','SUB_ADDRESS',
                'ORDER_CLASS','Y','X','ORDER_VOLUME','ORDER_WEIGHT',
                'BOX_NUM','ITEM_TYPE','ITEM_NM','ITEM_COUNT','ITEM_WEIGHT',
                'ITEM_COST','LOC_CUSTOM_CD','UNLOADING_TYPE','ORDER_TIME','S_ORDER_TIME',
                'E_ORDER_TIME','FORBIDDEN_TIME','OLD_CAR_NUM','OLD_VISIT_ORDER']

CUSTOM_VALUE_WITH_HEADER = {
    '운송타입': {
        '0': '일반',
        '1': '회수'
    },
    '주문유형': {
        '0': '일반',
        '1': '냉장'
    },
    '하차유형': {
        '0': '수작업',
        '1': '지게차'
    }
}

VAL_RANGE_WITH_HEADER = {
    # 대한민국 지도에 관한 일반정보의 경도범위는 124 – 132, 위도범위는 33 – 43 이다.
    '위도(Y)': [20.0, 55.0],
    '경도(X)': [110.0, 140.0]
}

COLS_TO_EMPTY = ['운송부피(CBM)','운송중량(kg)','박스수량','회피시간', '하차유형']

COLS_TO_REMOVE_ZERO = ['개별중량(kg)','상품수량', '개별원가','점착시간','점착시작시간','점착종료시간']

COLS_TO_REMOVE_ABNORMAL_HISTORY = ['배송차량', '배송순서']

COLS_TO_REMOVE_QUOTATION_MARK = ['배송지 주소', '배송지 상세주소']



def integrate_kt_and_mns_order(_kt_file_path, _mns_file_path, _kt_quotechar="'", _mns_quotechar='"'):
    # KT and MNS integration part


    kt_order_df = pd.read_csv(_kt_file_path, quotechar = _kt_quotechar, dtype = 'str')

    mns_order_df = pd.read_csv(_mns_file_path, quotechar = _mns_quotechar, dtype = 'str')

    # Remove comma in address
    kt_address_list = kt_order_df['ADDRESS'].tolist()
    kt_sub_address_list = kt_order_df['SUB_ADDRESS'].tolist()

    for idx, elem in enumerate(kt_address_list):
        elem = elem.replace(' ,', ' ')
        elem = elem.replace(', ', ' ')
        elem = elem.replace(',', ' ')
        elem = elem.replace('"','')
        kt_address_list[idx] = elem.replace("'",'')
        print(elem)
    for idx, elem in enumerate(kt_sub_address_list):
        elem = elem.replace(' ,', ' ')
        elem = elem.replace(', ', ' ')
        elem = elem.replace(',', ' ')
        elem = elem.replace('"','')
        kt_sub_address_list[idx] = elem.replace("'",'')
        print(elem)

    mns_address_list = mns_order_df['ADDRESS'].tolist()
    mns_sub_address_list = mns_order_df['SUB_ADDRESS'].tolist()

    for idx, elem in enumerate(mns_address_list):
        elem = elem.replace(' ,', ' ')
        elem = elem.replace(', ', ' ')
        elem = elem.replace(',', ' ')
        elem = elem.replace('"','')
        mns_address_list[idx] = elem.replace("'",'')
        print(elem)
    for idx, elem in enumerate(mns_sub_address_list):
        elem = elem.replace(' ,', ' ')
        elem = elem.replace(', ', ' ')
        elem = elem.replace(',', ' ')
        elem = elem.replace('"','')
        mns_sub_address_list[idx] = elem.replace("",'')
        print(elem)

    kt_order_df['ADDRESS'] = kt_address_list
    kt_order_df['SUB_ADDRESS'] = kt_sub_address_list
    mns_order_df['ADDRESS'] = mns_address_list
    mns_order_df['SUB_ADDRESS'] = mns_sub_address_list

    # Set time info as 0 - order time, forbidden time
    kt_order_df[["ORDER_TIME","S_ORDER_TIME","E_ORDER_TIME","FORBIDDEN_TIME"]] = 0
    mns_order_df[["ORDER_TIME","S_ORDER_TIME","E_ORDER_TIME","FORBIDDEN_TIME"]] = 0

    # Get and align column order
    kt_order_cols = kt_order_df.columns
    mns_order_cols = mns_order_df.columns

    mns_order_df = mns_order_df[kt_order_cols]

    # Concat two order dfs - KT and MNS
    concatenated_order_df = pd.concat([kt_order_df, mns_order_df])

    return [concatenated_order_df, [kt_order_df, mns_order_df]]



def convert_etl_format_to_excel_format(_etl_df):
    excel_converted_df = None
    tmp_list_size = len(_etl_df.index)

    for col_idx in range(0, len(EXCEL_HEADER)):
        col_to_pick = ERP_HEADER[col_idx]
        col_to_insert = EXCEL_HEADER[col_idx]

        list_to_insert = _etl_df[col_to_pick].tolist()

        if col_to_insert in CUSTOM_VALUE_WITH_HEADER.keys():
            tmp_val_dict = CUSTOM_VALUE_WITH_HEADER[col_to_insert]
            for idx, elem in enumerate(list_to_insert):
                if str(elem) in tmp_val_dict.keys():
                    list_to_insert[idx] = tmp_val_dict[str(list_to_insert[idx])]
        
        if col_to_insert in COLS_TO_EMPTY:
            list_to_insert = [''] * tmp_list_size

        if col_to_insert in VAL_RANGE_WITH_HEADER.keys():
            min_val = float(min(list(VAL_RANGE_WITH_HEADER[col_to_insert])))
            max_val = float(max(list(VAL_RANGE_WITH_HEADER[col_to_insert])))

            for idx, elem in enumerate(list_to_insert):
                try:
                    tmp_val = float(elem)
                except ValueError:
                    # Handle the exception
                    list_to_insert[idx] = ''
                    continue

                if math.isnan(tmp_val):
                    continue

                if tmp_val > max_val or tmp_val < min_val:
                    list_to_insert[idx] = ''
                else:
                    pass
        
        if col_to_insert in COLS_TO_REMOVE_ZERO:
            for idx, elem in enumerate(list_to_insert):
                if elem == 0 or elem == '0' or elem == '0.00':
                    list_to_insert[idx] = ''
        
        if col_to_insert in COLS_TO_REMOVE_QUOTATION_MARK:
            for idx, elem in enumerate(list_to_insert):
                elem = elem.replace('"', '')
                elem = elem.replace("'", '')
                list_to_insert[idx] = elem
    

        if excel_converted_df is None:
            excel_converted_df = pd.DataFrame({col_to_insert: list_to_insert})
            pass
        else:
            #print(col_to_insert)
            excel_converted_df[col_to_insert] = list_to_insert

    list_old_vehicle_num = _etl_df['OLD_CAR_NUM'].tolist()
    list_old_vehicle_visit_idx = _etl_df['OLD_VISIT_ORDER'].tolist()

    for idx in range(0, len(list_old_vehicle_visit_idx)):
        #print(type(list_old_vehicle_num[idx]), type(list_old_vehicle_visit_idx[idx]), list_old_vehicle_num[idx], list_old_vehicle_visit_idx[idx])

        if list_old_vehicle_num[idx] == '0' or float(list_old_vehicle_num[idx]) == 0:
            list_old_vehicle_num[idx] = ''
        if list_old_vehicle_visit_idx[idx] == '0' or float(list_old_vehicle_visit_idx[idx]) == 0:
            list_old_vehicle_visit_idx[idx] = ''

        '''
        if list_old_vehicle_visit_idx[idx] == '0' or float(list_old_vehicle_visit_idx[idx]) == 0:
            if list_old_vehicle_num[idx] == '0':
                list_old_vehicle_num[idx] = ''
                list_old_vehicle_visit_idx[idx] = ''
        '''
    
    excel_converted_df['배송차량'] = list_old_vehicle_num
    excel_converted_df['배송순서'] = list_old_vehicle_visit_idx

    for idx in range(0, len(list_old_vehicle_visit_idx)):
        if list_old_vehicle_visit_idx[idx] == '0':
            list_old_vehicle_visit_idx[idx] = ''

    excel_converted_df = excel_converted_df[EXCEL_HEADER]

    return excel_converted_df


def write_order_to_excel(_df_to_write, _result_file_path, _template_file_path='./template/delivery_order_template.xlsx'):
    # Get excel template
    

    # Align columns to fit excel format (Result: 'converted_etl_df')


    # Copy template and write to copied template (=result file)
    shutil.copy(_template_file_path, _result_file_path)

    writer = pd.ExcelWriter(_result_file_path, engine='openpyxl', mode='w')
    writer.book = load_workbook(_result_file_path)
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

    _df_to_write.to_excel(writer, sheet_name='운송주문입력', startrow=7, startcol=0, header=False, index=False, )
    writer.save()
    writer.close()
    

if __name__ == "__main__":

    mid_result_dir = './mid_result/'

    center_id = 1031
    date = 20210323

    kt_quotechar = "'"
    mns_quotechar = "'"

    separated_file_dir = './order_data/{0}/raw/'.format(date)

    for i in [1, 2, 3, 4, 5, 6]:
        print("Processing order data of cluster {0}".format(i))

        kt_file_name = '{0}_{1:02d}_{2}_KT.csv'.format(center_id, i, date)
        mns_file_name = '{0}_{1:02d}_{2}_MnS.csv'.format(center_id, i, date)
        concatenated_input_file_name = '{0}_{1:02d}_{2}.csv'.format(center_id, i, date)
        integrated_file_name = '{0}_{1:02d}_{2}_integrated.csv'.format(center_id, i, date)

        integrated_output_file_name = '{0}_{1:02d}_{2}_out.xlsx'.format(center_id, i, date)

        result_dir = './result/{0}/'.format(date)

        kt_file_path = separated_file_dir + kt_file_name
        mns_file_path = separated_file_dir + mns_file_name

        integrated_output_file_path = result_dir + integrated_output_file_name
        concatenated_input_file_path = separated_file_dir + concatenated_input_file_name

        # Get file path as single string - after, use QT

        concatenated_order_df, _ = integrate_kt_and_mns_order(kt_file_path, mns_file_path, kt_quotechar, mns_quotechar)

        #concatenated_order_df = pd.read_csv(concatenated_input_file_path)

        # Save to excel
        excel_df = convert_etl_format_to_excel_format(concatenated_order_df)

        write_order_to_excel(excel_df, integrated_output_file_path)

        print("Get integrated order data of cluster {0}".format(i))

    print("FINISH!!!!!!")