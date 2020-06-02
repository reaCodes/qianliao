import os

import numpy as np
import xlwings as xw


def dir_list(path, all_file, all_file_name):
    file_list = os.listdir(path)

    for file_name in file_list:
        file_path = os.path.join(path, file_name)
        if os.path.isdir(file_path):
            dir_list(file_path, all_file, all_file_name)
        else:
            all_file.append(file_path)
            all_file_name.append(file_name)

    return all_file, all_file_name


def show_files_name():
    # wb = xw.Book.caller()
    wb = xw.Book("qianliao_config.xlsm")
    st = wb.sheets[0]
    start_path = st.range("I2").value

    files_full_path = list()
    files_only_name = list()
    files_path, files_name = dir_list(start_path, files_full_path, files_only_name)

    cell_range = wb.app.selection
    row_num = cell_range.row
    # col_num = cell_range.column
    st.range('C:C').clear_contents()
    st.range('H:H').clear_contents()
    st.range(row_num, 8).value = np.array(files_path).reshape(len(files_path), 1)
    st.range(row_num, 3).value = np.array(files_name).reshape(len(files_name), 1)


if __name__ == "__main__":
    # dir_path = "C:\\Users\\zhong\\Documents\\Study\\program_x"
    # all_files = list()
    # all_files = dir_list(dir_path, all_files)
    # print(all_files)
    show_files_name()
