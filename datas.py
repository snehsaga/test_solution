import json
import xlrd


def process_data(path):
    """Writes json data to a .json file from excel sheet

    :param path: string representing path where excel sheet is located
    """
    sheet = xlrd.open_workbook(path).sheet_by_name('MICs List by CC')
    excel_data = []
    dict_keys = list(map(lambda obj: str(obj.encode('utf-8')), sheet.row_values(0)))
    rows = sheet.nrows
    for row in range(1, rows):
        data = list(map(lambda obj: str(obj.encode('utf-8')), sheet.row_values(row)))
        data_dict = {}
        for dict_key, value in zip(dict_keys, data):
            data_dict[dict_key] = value
        excel_data.append(data_dict)
    with open('result.json', 'w') as file:
        json.dump(excel_data, file)


if __name__ == '__main__':
    process_data('ISO10383_MIC.xls')
