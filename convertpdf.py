import logging
import os.path
import pathlib
import sys
import time
from logging.handlers import TimedRotatingFileHandler

import pandas as pd
import PySimpleGUI as sg
import yaml

working_dir = str(pathlib.Path().absolute())

pdf_path = os.path.join(working_dir, "sample_pdf_data", "01 09 2021 to 28 09 2021 JobAttendanceDetailedReport-1.pdf")
# https://blog.alivate.com.au/poppler-windows/
poppler_path = os.path.join(working_dir, "poppler-0.68.0", "bin", "pdftotext.exe")

version = 'v1.1 - 27/10/2021'
copy_right = 'Developed by Sigma2k Technology Pte Ltd'


logger = logging.getLogger(__name__)
_LOG_FOLDER = os.path.join(working_dir, 'logs')
_LOG_FILE = os.path.join(_LOG_FOLDER, 'result.log')
_FMT = '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)s] - %(message)s'
_FORMATTER = logging.Formatter(_FMT)
_STREAM_HANDLER = logging.StreamHandler(sys.stdout)
_STREAM_HANDLER.setFormatter(_FORMATTER)
logger.setLevel(logging.DEBUG)
logger.addHandler(_STREAM_HANDLER)

try:
    os.makedirs(_LOG_FOLDER)
    print(('Created log folder at path {}'.format(_LOG_FOLDER)), flush=True)
except OSError as e:
    if os.path.exists(_LOG_FOLDER):
        print('Log folder already exist.', flush=True)
    else:
        print(('Failed to create log folder at path {}, reason: {}'.format(_LOG_FOLDER, e)), flush=True)

_FILE_HANDLER = TimedRotatingFileHandler(_LOG_FILE, when='midnight')
_FILE_HANDLER.setFormatter(_FORMATTER)
logger.addHandler(_FILE_HANDLER)


def poppler_pdf_to_text(pdf_file=pdf_path, result_folder=working_dir):
    file_name = os.path.basename(pdf_file).split('.')[0]
    raw_file = os.path.join(working_dir, 'result', 'raw', f'{file_name}_raw_poppler.txt')
    raw_des_file = os.path.join(working_dir, 'result', 'raw', f"{file_name}_description.txt")
    raw_table_file = os.path.join(working_dir, 'result', 'raw', f"{file_name}_table.txt")
    raw_csv_file = os.path.join(working_dir, 'result', 'raw', f"{file_name}_csv.txt")
    cmd = f'{poppler_path} "{pdf_file}" "{raw_file}" -layout'
    os.system(cmd)

    # This Step make clean the data (remove unnecessary lines)
    # And Store 2 txt files: description and table
    line_need_remove = []
    end_description = 0
    # first_header = False
    # first_ks_day = False
    task_name = ''  # Store Current TASK of CURRENT SHIFT
    with open(raw_file) as f:
        data = f.readlines()
    for i in range(len(data)):
        line = data[i]
        if 'ID NUMBER' in line:
            line = line.replace('ID NUMBER', 'ID_NUMBER')
            line = line.replace('', '')
            # line = line.replace('ACTUAL', 'ACTUAL_HOURS')
            # line = line.replace('ACCOUNTABLE', 'ACCOUNTABLE_HOURS')
            # line = line.replace("WORKERS'", "WORKERS'")
            data[i] = line
            # if not first_header:
            #     first_header = True
            # else:
            #     line_need_remove.append(i)
        if 'Subcontractor' in line and 'SIGNATORIES:' in line:
            line = line.replace('Subcontractor', '').replace('SIGNATORIES:', '').strip()
            data[i] = line
        elif 'Yard Supervisor' in line:
            line = line.replace('Yard Supervisor', '')
            data[i] = line
        elif 'Yard Foreman' in line:
            line = line.replace('Yard Foreman', '')
            data[i] = line
        elif 'Yard Superintendent' in line:
            line = line.replace('Yard Superintendent', '')
            data[i] = line
        elif 'TRADE' in line and 'HOURS' in line:
            line_need_remove.append(i)
        # elif line.count(':') == 2:
        #     line_need_remove.append(i)
        elif line.isspace():
            line_need_remove.append(i)
        elif 'SUMMARY FOR SHIFT' in line:
            line_need_remove.append(i)
        elif 'Page' in line and 'of' in line:
            line_need_remove.append(i)
        # elif 'KS-DAY' in line:
        #     if not first_ks_day:
        #         first_ks_day = True
        #     else:
        #         line_need_remove.append(i)
        elif 'SUBTOTAL FOR DAY SHIFT' in line or 'SUBTOTAL FOR NIGHT SHIFT' in line or 'SUMMARY FOR JOB CARD' in line or 'ATTENDANCE' in line or 'END OF REPORT' in line:
            line_need_remove.append(i)
        if 'DATE' in line and 'JOB CARD' not in line:
            line_need_remove.append(i)
            # if not end_description:
            #     end_description = i
    for i in line_need_remove[::-1]:
        data.pop(i)

    for i in range(len(data)):
        line = data[i]
        if 'SHIFT A' in line:
            if not end_description:
                end_description = i
                task_name = line
    out_description_txt = []
    for i in range(end_description, -1, -1):
        out_description_txt.insert(0, data[i])
        data.pop(i)

    count = 0
    for line in data:
        if 'G' == line[0]:
            count += 1
    logger.debug(f'Total Job {count}')

    with open(raw_des_file, 'w') as f:
        f.writelines(out_description_txt)

    with open(raw_table_file, 'w') as f:
        f.writelines(data)

    # generate CSV file
    headers = ['ID_NUMBER', 'NAME', 'NATIONALITY', "WORKERS'", 'SKILL', 'START', 'END', 'MEALS', 'REMARKS', 'ACTUAL',
               'ACCOUNTABLE']
    h_index = []    # [0, 18, 54, 66, 91, 110, 129, 142, 150, 165, 172]
    line_need_remove = []
    first_header = False
    header_str = ''
    task_list = []  # For insert Task of SHIFT
    for i in range(len(data)):
        line = data[i]
        if 'SHIFT A' in line:
            task_name = line
            line_need_remove.append(i)
            continue
        elif 'ID_NUMBER' in line:
            if not first_header:
                first_header = True
            else:
                line_need_remove.append(i)
            h_index = []
            tmp_line = line
            for h in headers:
                if len(line) < 150:
                    if h == "WORKERS'":
                        h_index.append(line.index(h) - 2)
                    elif h in ['ACTUAL', 'ACCOUNTABLE']:
                        h_index.append(line.index(h) + 2)
                    else:
                        h_index.append(line.index(h))
                else:
                    if h == "WORKERS'":
                        h_index.append(line.index(h) - 7)
                    elif h in ['ACTUAL', 'ACCOUNTABLE']:
                        h_index.append(line.index(h) + 7)
                    else:
                        h_index.append(line.index(h))
                if h != 'ACCOUNTABLE':
                    h_idx = tmp_line.index(h) + len(h)
                    tmp_line = tmp_line[:h_idx] + ',' + tmp_line[h_idx:]
            header_str = tmp_line
        elif line.count(':') != 2:
            # get Start and End Time
            next_line = data[i+1]
            idx_name = headers.index('NAME')
            idx_national = headers.index('NATIONALITY')
            str_name = next_line[h_index[idx_name]:h_index[idx_national]].strip()   # 01 10 2021.pdf
            # if 'THIRUNAVUKKARASU' in str_name:
            #     print('debug')

            idx_start_time = headers.index('START')
            idx_end_time = headers.index('END')
            str_start_time = next_line[h_index[idx_start_time]:h_index[idx_end_time]]
            str_end_time = next_line[h_index[idx_end_time]:h_index[idx_end_time+1]]
            line_need_remove.append(i+1)

            task_list.append(task_name.replace('\n', '').strip())
            length = len(line)
            new_line = ''
            for idx in range(len(h_index)):
                current_index = h_index[idx]
                if idx == len(h_index) - 1:
                    next_index = length
                else:
                    next_index = h_index[idx + 1]
                item_data = line[current_index:next_index]
                if idx != len(h_index) - 1:
                    # logger.debug(len(item_data))
                    # is_space = item_data.isspace()
                    # if is_space:
                    #     item_data = item_data[::-1].replace('  ', 'N,', 1)[::-1]
                    # else:
                    add_data = ','
                    if idx == idx_start_time:
                        add_data = f'{add_data}{str_start_time.strip()[::-1]} '
                    elif idx == idx_end_time:
                        add_data = f'{add_data}{str_end_time.strip()[::-1]} '
                    elif idx == idx_name:
                        add_data = f'{add_data} {str_name.strip()[::-1]} '
                    item_data = f"{item_data.strip()} "
                    item_data = item_data[::-1].replace(' ', add_data, 1)[::-1]
                    # logger.debug(len(item_data))
                # logger.debug(item_data)
                new_line += item_data
            # logger.debug(new_line)
            data[i] = new_line

    for i in line_need_remove[::-1]:
        data.pop(i)

    data[0] = header_str.replace("WORKERS'", "WORKERS'TRADE").replace('ACTUAL', 'ACTUAL HOURS').replace('ACCOUNTABLE',
                                                                                                        'ACCOUNTABLE HOURS')
    with open(raw_csv_file, 'w') as f:
        f.writelines(data)

    # Parse DESCRIPTION, project
    des_data = {}
    for i in range(len(out_description_txt)):
        line = out_description_txt[i]
        if 'BU' in line:
            data = line.split('BU')[1].lstrip(':').strip()
            des_data.update({'BU': data})
        elif 'PROJECT NAME' in line:
            data = line.split('PROJECT NAME')[1].lstrip(':').strip()
            des_data.update({'PROJECT NAME': data})
        elif 'DESCRIPTION' in line:
            data = line.split('DESCRIPTION')[1].replace('\n', '').lstrip(':').strip()
            for j in range(i + 1, len(out_description_txt)):
                txt = out_description_txt[j]
                if ' ' == txt[0]:
                    data += ' ' + txt.replace("\n", "").strip()
                else:
                    break
            # logger.debug(data)
            des_data.update({'DESCRIPTION': data})

    df = pd.read_csv(raw_csv_file, sep=' *, *')
    idx = 2
    for key, value in des_data.items():
        df.insert(idx, key, [value] * df['ID_NUMBER'].size)
        idx += 1
    df.insert(4, 'TASK', task_list)

    df['START'] = pd.to_datetime(df['START'], format='%d/%m/%Y %H:%M').dt.strftime('%m/%d/%Y %H:%M')
    df['END'] = pd.to_datetime(df['END'], format='%d/%m/%Y %H:%M').dt.strftime('%m/%d/%Y %H:%M')
    headers = ['ID_NUMBER', 'NAME', 'START', 'END']

    df.to_csv(os.path.join(result_folder, f"{file_name}.csv"), columns=headers, index=False)

    # # Create a Pandas Excel writer using XlsxWriter as the engine.
    # writer = pd.ExcelWriter(os.path.join(result_folder, f'{file_name}.xlsx'), engine='xlsxwriter')
    #
    # # Convert the dataframe to an XlsxWriter Excel object.
    # df.to_excel(writer, sheet_name='Sheet1', columns=headers, index=False)
    #
    # # Close the Pandas Excel writer and output the Excel file.
    # writer.save()
    logger.debug(df)

class MainGUI:
    icon_main = os.path.join(working_dir, 'icon.ico')

    KEY_THREAD_WORKER_END = '--scan_api_thread--'

    def __init__(self):
        self._gui_setting_path = os.path.join(working_dir, '_gui_setting.yaml')
        self._default_gui_setting = self.generator_default_gui_setting()
        self._current_gui_setting = None
        self.update_gui_setting()

    @staticmethod
    def generator_default_gui_setting():
        default = dict()
        default['output_folder'] = working_dir

        return default

    @property
    def current_gui_setting(self):
        return self._current_gui_setting

    def update_gui_setting(self):
        self._current_gui_setting = self._default_gui_setting

        try:
            if not os.path.exists(self._gui_setting_path):
                with open(self._gui_setting_path, 'a') as f:
                    yaml.dump(self._default_gui_setting, f, indent=2, sort_keys=False)
                    f.close()
            else:
                with open(self._gui_setting_path, 'r') as f:
                    tmp_setting = yaml.load(f, Loader=yaml.FullLoader)
                    f.close()
                for k, v in self._current_gui_setting.items():
                    setting_value = tmp_setting.get(k, None)
                    if setting_value is not None and setting_value != v:
                        self._current_gui_setting[k] = setting_value
                logger.debug('Updated gui setting!')
        except Exception as e:
            logger.debug(f'update_gui_setting Error: {e}')
            import traceback
            traceback.print_exc(limit=5)

    def save_gui_setting(self):
        ret = True
        with open(self._gui_setting_path, 'w+', encoding='utf-8') as f:
            try:
                yaml.dump(self._current_gui_setting, f, indent=2, sort_keys=False)
            except Exception as e:
                logger.debug(e)
                ret = False
            f.close()
        return ret

    def init_main_gui(self):
        layout = [
            [sg.FilesBrowse(button_text='PDF Files', size=(15, 1)), sg.InputText(k='PDF_FILE', size=(80, 1))],
            [sg.FolderBrowse(button_text='Output Folder', size=(15, 1)),
             sg.InputText(default_text=self.current_gui_setting['output_folder'], k='OUT_FOLDER', size=(80, 1))],
            [sg.B('Convert')]
        ]
        window = sg.Window(f"PDF Convert Tool {version} - {copy_right}", layout=layout, finalize=True,
                           icon=MainGUI.icon_main)

        last_save_setting_time = time.time()
        while True:
            event, value = window.read(100)
            if event in ['Exit', sg.WIN_CLOSED]:
                break
            elif event is sg.TIMEOUT_EVENT:
                if time.time() - last_save_setting_time > self.current_gui_setting.get('interval_save_setting_time', 2):
                    self.current_gui_setting['output_folder'] = value['OUT_FOLDER']
                    self.save_gui_setting()
                    last_save_setting_time = time.time()
            elif event == 'Convert':
                pdf_files = value['PDF_FILE'].split(';') if value['PDF_FILE'] else []
                result_folder = value['OUT_FOLDER']
                if not os.path.isdir(result_folder):
                    sg.popup_error(f'Not Folder\n{result_folder}')
                    continue
                for file in pdf_files:
                    if not os.path.isfile(file):
                        sg.popup_error(f'Not File\n\n{file}\n\nContinue with other files...')
                        continue
                    poppler_pdf_to_text(file, result_folder)
                sg.popup_ok('Complete')
        window.close()


if __name__ == '__main__':
    MainGUI().init_main_gui()

