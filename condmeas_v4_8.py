import pyvisa
import os
import datetime
import shutil
import time
import numpy as np
from threading import Thread

# Keithleyの識別子一覧
KEITHLEY_MAP = {
    '1': '0x05E6::0x2450::04588911',
    '3': '0x05E6::0x2450::04048286',
    '4': '0x05E6::0x2450::04131783',
    '6': '0x05E6::0x2450::04131618'
}

# Data_Directory一覧
DATA_DIRECTORY_MAP = {
    '1': r"E:\cond_data",
    '3': r'C:\Data',
    '4': r'C:\Data',
    '6': r'C:\Users\user\Documents\data'
}

#数値への変換判定、うまくいかないならFalseを返す
def validate_number(number, target_type):
    try:
        target_type(number) # 渡された型（intやfloat）で直接キャストを試みる
        return True
    except ValueError:
        return False

#プログラム起動時処理
def initialize_measurement():
    # インスタンス生成、おまじないのようなもの
    rm = pyvisa.ResourceManager()
    # PCに接続された計測器のVISAリソース名の取得
    visa_list = rm.list_resources()
    # Keithleyを指定
    keithley_path, furnace_number = find_keithley(rm, visa_list)
    # ResourceManagerを閉じる
    rm.close()
    
    return keithley_path, furnace_number

# keithley探し
def find_keithley(rm, visa_list):
    connected_keithleys = {}
    
    for machine_name in visa_list:
        for f_num, identifier in KEITHLEY_MAP.items():
            if identifier in machine_name:
                connected_keithleys[f_num] = machine_name

    if len(connected_keithleys) == 0:
        print('Keithleyをつないでやり直してください。\nEnterを押すと終了します。')
        print('Please connect keithley and start over.\nPress Enter to exit.')
        rm.close()
        input()
        exit()
    elif len(connected_keithleys) == 1:
        furnace_number = list(connected_keithleys.keys())[0]
        find_name = connected_keithleys[furnace_number]
    else:
        print('複数のKeithleyが接続されています。どれを使用するか選択してください。')
        print("Multiple Keithley units are connected. Please select which one to use.")
        print('Unit:', ', '.join(connected_keithleys.keys()))
        while True:
            furnace_number = input('使用する号機を入力してください (例: 1): ')
            if furnace_number in connected_keithleys:
                find_name = connected_keithleys[furnace_number]
                break
            print('入力された号機のKeithleyは検出されていません。正しい号機を入力してください。')
            print('The Keithley unit with the entered serial number was not found. Please enter the correct serial number.')

    inst_1 = rm.open_resource(find_name)
    verify_connection(rm, inst_1)
    inst_1.close()
    
    return find_name, furnace_number

# keithleyのエラーチェック
def verify_connection(rm, inst):
    out = inst.query('*IDN?')

    if 'KEITHLEY INSTRUMENTS' not in out:
        print('keithleyを再起動してやり直してください。\nEnterを押すと終了します。')
        print('Please restart keithley and try again.\nPress Enter to exit.')
        inst.close()
        rm.close()
        x = input()
        exit()
    
    inst.write(':ABORt')
    out = inst.query(':SYSTem:ERRor?')
    s_comma = out.split(',')
    if s_comma[0] != "0":
        print('keithleyを再起動してやり直してください。\nEnterを押すと終了します。')
        print('Please restart keithley and try again.\nPress Enter to exit.')
        inst.close()
        rm.close()
        x = input()
        exit()
    
    inst.write('*CLS;*RST')

# データ置き場の設定
def select_save_directory(data_path):
    files_dir = [f for f in os.listdir(datapath) if os.path.isdir(os.path.join(datapath, f))]
    file_maxnumber = len(files_dir) - 1
    
    while True:
        print('どのフォルダにデータを保存するか選択してください。\n新たにフォルダを作る場合は「new」と入力してください。')
        print('Select which folder you want to save the data in.\nTo create a new folder, type "new".')
        for i, name in enumerate(files_dir):
            print(i, name)
        file_number = input('回答 answer:')
        
        # 新しいディレクトリを作る場合はやり直すようにする
        if file_number == 'new':
            print('プログラムを終了し、フォルダを作成してから再度やり直してください。\nEnterを押すと終了します。')
            print('Exit the program, create the folder, and try again.\nPress Enter to exit.')
            x = input()
            exit()
            
        # 整数かどうかの確認
        if not validate_number(file_number, int):
            print('正しい数値を入力してください。')
            print('Please enter the correct value.')
            continue
            
        file_number = int(file_number)
        
        # 範囲内かどうかの確認
        if not (0 <= file_number <= file_maxnumber):
            print('正しい数値を入力してください。')
            print('Please enter the correct value.')
            continue
            
        break

    data_path = fr"{data_path}\{files_dir[file_number]}"    
    return data_path

# データ保存場所の作成
def create_save_folder(data_path, sample_name, sample_comment, furnace_number):
    time_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

    if sample_comment == '':
        filename = f"{time_now}_{sample_name}_{furnace_number}"
    else:
        filename = f"{time_now}_{sample_name}_{sample_comment}_{furnace_number}"
    data_path = fr"{data_path}\{filename}"

    os.makedirs(data_path, exist_ok=True)
    shutil.copy(fr"condmeas_setfile\measure_excel.xlsm", fr"{data_path}\{filename}.xlsm")
    shutil.copy(fr"condmeas_setfile\measurement_condition.txt", fr"{data_path}\measurement_condition.txt")
    shutil.copy(fr"condmeas_setfile\setting_logger_files_{furnace_number}.txt", fr"{data_path}\setting_logger_files.txt")
    return data_path

def get_sample_float_value(prompt_ja, prompt_en, input_prompt):
    print(prompt_ja)
    print(prompt_en)
    while True:
        val_str = input(input_prompt)
        if validate_number(val_str, float):
            if float(val_str) > 0:
                return float(val_str)
        print('正しい数値を入力してください。')
        print('Please enter the correct value.')

# サンプルデータリストの作成、重さ、高さ、直径、端子間距離
def input_sample_info():
    sample_weight = get_sample_float_value('サンプル重さ[g]', 'Sample weight [g]', 'weight[g]:')
    sample_height = get_sample_float_value('サンプル高さ[mm]', 'Sample height [mm]', 'height[mm]:')
    sample_diameter = get_sample_float_value('サンプル直径[mm]', 'Sample diameter [mm]', 'diameter[mm]:')
    sample_distance = get_sample_float_value('端子間距離[mm]', 'Terminal distance [mm]', 'Terminal distance[mm]:')
    return [sample_weight, sample_height, sample_diameter, sample_distance]

# 測定条件の決定、スタート電圧、ストップ電圧、ディレイタイム、ステップ数、測定レンジ、電流リミット、測定スピード、接続方式、間隔
def load_measurement_conditions(data_path):
    with open( fr"{data_path}\measurement_condition.txt", mode = 'r') as f:
        condition = f.readlines()
        start_voltage = float(condition[0].rstrip('\n'))
        end_voltage = float(condition[1].rstrip('\n'))
        deley_time = float(condition[2].rstrip('\n'))
        step_number = int(condition[3].rstrip('\n'))
        current_range = condition[4].rstrip('\n')
        current_limit = float(condition[5].rstrip('\n'))
        measurement_speed = float(condition[6].rstrip('\n'))
        connection_method = condition[7].rstrip('\n')
        repeat_time = int(condition[8].rstrip('\n'))
    
    print('測定条件')
    print('measurement condition')
    print(f'スタート電圧：{start_voltage}')
    print(f'ストップ電圧：{end_voltage}')
    print(f'ディレイ時間：{deley_time}')
    print(f'ステップ数：{step_number}')
    print(f'測定レンジ：{current_range}')
    print(f'電流リミット：{current_limit}')
    print(f'測定スピード：{measurement_speed}')
    print(f'接続方式：{connection_method}')
    print(f'間隔(s):{repeat_time}')
    
    while True:
        print('以上の条件でいいですか？')
        print('Are the above conditions correct?')
        out = input('y/n:')
        if out == 'y':
            break
        elif out == 'n':
            print('正しい値を入力してください。')
            print('Please enter the correct value.')
            # スタートストップ電圧
            while True:
                while True:
                    start_voltage_str = input('スタート電圧：')
                    if validate_number(start_voltage_str, float):
                        if -20 <= float(start_voltage_str) <= 20:
                            start_voltage = float(start_voltage_str)
                            break
                    print('正しい数値を入力してください。')
                    print('Please enter the correct value.')
                
                while True:
                    end_voltage_str = input('ストップ電圧：')
                    if validate_number(end_voltage_str, float):
                        if -20 <= float(end_voltage_str) <= 20:
                            end_voltage = float(end_voltage_str)
                            break
                    print('正しい数値を入力してください。')
                    print('Please enter the correct value.')
                
                if start_voltage < end_voltage: break
                else:
                    print('スタート電圧はストップ電圧より低い必要があります。')
                    print('The start voltage must be lower than the stop voltage.')
            
            # ディレイ
            while True:
                deley_time_str = input('ディレイ時間：')
                if validate_number(deley_time_str, float):
                    if 0.00005 <= float(deley_time_str) <= 10:
                        deley_time = float(deley_time_str)
                        break
                print('正しい数値を入力してください。')
                print('Please enter the correct value.')

            # ステップ数
            while True:
                step_number_str = input('ステップ数：')
                if validate_number(step_number_str, int):
                    if 2 <= int(step_number_str) <= 101:
                        step_number = int(step_number_str)
                        break
                print('正しい整数を入力してください。')
                print('Please enter the correct value.')
            
            # 測定レンジ
            while True:
                current_range = input('測定レンジ：')
                if current_range == 'Auto' or current_range == 'Fixed':
                    break
                print('測定レンジは「Auto」あるいは「Fixed」で回答してください。')
                print('Please answer with “Auto” or “Fixed” for the measurement range.')

            # 電流リミット
            while True:
                current_limit_str = input('電流リミット：')
                if validate_number(current_limit_str, float):
                    if 0.000001 <= float(current_limit_str) <= 1:
                        current_limit = float(current_limit_str)
                        break
                print('正しい数値を入力してください。')
                print('Please enter the correct value.')

            # 測定スピード
            while True:
                measurement_speed_str = input('測定スピード：')
                if validate_number(measurement_speed_str, float):
                    if 0.01 <= float(measurement_speed_str) <= 10:
                        measurement_speed = float(measurement_speed_str)
                        break
                print('正しい数値を入力してください。')
                print('Please enter the correct value.')

            # 接続方式
            while True:
                connection_method = input('接続方式：')
                if connection_method == 'Remote' or connection_method == 'Local':
                    break
                print('測定レンジは「Remote」あるいは「Local」で回答してください。')
                print('Please answer with “Remote” or “Local” for the measurement range.')
            
            # 間隔
            while True:
                repeat_time_str = input('測定間隔(s):')
                if validate_number(repeat_time_str, int):
                    if int(repeat_time_str) > 0:
                        repeat_time = int(repeat_time_str)
                        break
                print('正しい整数を入力してください。')
                print('Please enter the correct value.')

            with open( fr"{data_path}\measurement_condition.txt", mode = 'w') as f:
                f.write(f"{start_voltage}\n{end_voltage}\n{deley_time}\n{step_number}\n{current_range}\n{current_limit}\n{measurement_speed}\n{connection_method}\n{repeat_time}")

        else:
            print('y/n で回答してください。')
            print('Please answer in y/n.')

    condition_list = [start_voltage, end_voltage, deley_time, step_number, current_range, current_limit, measurement_speed, connection_method, repeat_time]
    return condition_list

# ロガーファイル読み込み
def load_logger_info(data_path, furnace_number):
    with open(fr"{data_path}\setting_logger_files.txt", mode = 'r') as f:
        logger_files = f.readlines()
        logger_directry = logger_files[0].rstrip('\n')
        logger_6gouki = logger_files[1].rstrip('\n')
        if furnace_number in ('1', '2', '6'):
            print('酸素ポンプを利用しますか？')
            print('Do you use pO2 controller?')
            while True:
                answer = input('y/n:')
                if answer == 'n':
                    channel_data = logger_files[1 + int(furnace_number)].rstrip('\n').split(',')
                    logger_channel = [channel_data[0], channel_data[1]]
                    break
                elif answer == 'y':
                    channel_data = logger_files[1 + int(furnace_number)].rstrip('\n').split(',')
                    logger_channel = [channel_data[0], channel_data[2]]
                    break
                else:
                    print('y/n で回答してください。')
                    print('Please answer in y/n.')
            
        elif furnace_number in ('3', '4'):
            logger_channel = [logger_files[1 + int(furnace_number)].rstrip('\n'), '0']
            
        else:
            print(f'{furnace_number}号機用のロガー設定が見つかりません。プログラムを確認してください。\nEnterを押すと終了します。')
            print(f'The logger settings for Unit {furnace_number} cannot be found. If this message appears, please check the program.\nPress Enter to exit.')
            x = input()
            exit()
    
    try:
        log_files = [os.path.join(logger_directry, f) for f in os.listdir(logger_directry) if os.path.isfile(os.path.join(logger_directry, f))]
    except:
        print('ロガーにアクセスできません。Windows資格情報をチェックしてください。\nEnterを押すと終了します。')
        print('Unable to access logger, please check Windows credentials.\nPress Enter to exit.')
        x = input()
        exit()
    else:
        log_files.sort(key=os.path.getmtime, reverse=True)
        log_file = log_files[0]
    
    print('ロガーファイルの場所は以下の通りで間違いないですか？')
    print('Is the path to the logger file as follows?')
    print(log_file)
    while True:
        answer = input('y/n:')
        if answer == 'y':
            break
        elif answer == 'n':
            while True:
                print('ロガーファイル名を入力してください。')
                print('Enter a logger file name.')
                log_file = input('file name:')
                log_file = os.path.join(logger_directry, log_file)
                print(log_file + 'にアクセスします。')
                try:
                    with open(log_file) as f:
                        pass
                except FileNotFoundError:
                    print('正しいファイル名を入力してください。')
                    print('Please enter the correct file name.')
                else:
                    print('このファイルでよろしいですか？')
                    print('Is this file correct?')
                    break
        else:
            print('y/n で回答してください。')
            print('Please answer in y/n.')

    
    if furnace_number == '6':
        try:
            log_files_6gouki = [os.path.join(logger_6gouki, f) for f in os.listdir(logger_6gouki) if os.path.isfile(os.path.join(logger_6gouki, f))]
        except:
            print('ロガーにアクセスできません。Windows資格情報をチェックしてください。\nEnterを押すと終了します。')
            print('Unable to access logger, please check Windows credentials.\nPress Enter to exit.')
            x = input()
            exit()
        else:
            log_files_6gouki.sort(key=os.path.getmtime, reverse=True)
            log_file_6gouki = log_files_6gouki[0]

        print('6号機のロガーファイルの場所は以下の通りで間違いないですか？')
        print('Is the path to the logger file of 6-gouki as follows?')
        print(log_file_6gouki)
        while True:
            answer = input('y/n:')
            if answer == 'y':
                break
            elif answer == 'n':
                while True:
                    print('ロガーファイル名を入力してください。')
                    print('Enter a logger file name.')
                    log_file_6gouki = input('file name:')
                    log_file_6gouki = os.path.join(logger_6gouki, log_file_6gouki)
                    print(log_file_6gouki + 'にアクセスします。')
                    try:
                        with open(log_file_6gouki) as f:
                            pass
                    except FileNotFoundError:
                        print('正しいファイル名を入力してください。')
                        print('Please enter the correct file name.')
                    else:
                        print('このファイルでよろしいですか？')
                        print('Is this file correct?')
                        break
            else:
                print('y/n で回答してください。')
                print('Please answer in y/n.')
    else:
        log_file_6gouki = ""
    
    logger_infoList = [log_file, logger_channel, log_file_6gouki]
    with open(fr"{data_path}\logger_files.txt", mode = 'w') as f:
        f.write(log_file + '\n')
        f.write(log_file_6gouki + '\n')
        f.write(logger_channel[0] + ',' + logger_channel[1] + '\n')
    return logger_infoList

# 測定ルーチン
def measure_conductivity(nowtime, measurement_number, keithley_path, data_path, furnace_number, measurement_condition_List, logger_infoList, sample_info):

    global has_error

    print('測定中…終了しないでください。')
    print('During measurement...do not exit.')

    try:
        if furnace_number == '6':
            with open(logger_infoList[2], mode = 'r') as f:
                log_List = f.readlines()
                temp = log_List[len(log_List) - 1].rstrip('\n').split(',')[6]
        else:
            with open(logger_infoList[0], mode = 'r') as f:
                log_List = f.readlines()
                temp = log_List[len(log_List) - 1].rstrip('\n').split(',')[int(logger_infoList[1][0]) + 2].lstrip().lstrip('+')
        
        if logger_infoList[1][1] == '0':
            pO2 = ''
        else:
            with open(logger_infoList[0], mode = 'r') as f:
                log_List = f.readlines()
                Log_pO2 = log_List[len(log_List) - 1].rstrip('\n').split(',')[int(logger_infoList[1][1]) + 2].lstrip().lstrip('+')
                pO2 = str(10 ** (-1 * float(Log_pO2)))
        
        # インスタンス生成、おまじないのようなもの
        rm = pyvisa.ResourceManager()
        # ケースレーとの通信を開く
        keithley = rm.open_resource(keithley_path)
        # 接続テスト
        verify_connection(rm, keithley)
        # 電流を測定するようにセット
        keithley.write(':SENSe:FUNCtion "CURRent" ')
        # 電流の上限、下限
        if measurement_condition_List[4] == 'Auto':
            sCmds = ':SENSe:CURRent:RANGe:AUTO ON'
        else:
            current_range = '{:E}'.format(float(measurement_condition_List[5]))
            sCmds = ':SENSe:CURRent:RANGe ' + current_range
        keithley.write(sCmds)
        # 4端子か2端子か
        if measurement_condition_List[7] == 'Remote':
            sCmds = ':SENSe:CURRent:RSENse ON'
        else:
            sCmds = ':SENSe:CURRent:RSENse OFF'
        keithley.write(sCmds)
        # 測定スピード
        keithley.write(':SENSe:CURRent:NPLCycles ' + str(measurement_condition_List[6]))
        # 電圧による測定
        keithley.write(':SOURce:FUNCtion VOLTage')
        # 測定電圧レンジ
        Voltage_range = str(measurement_condition_List[1] - measurement_condition_List[0])
        keithley.write(':SOURce:VOLTage:RANGe ' + Voltage_range)
        # 測定電圧レンジに対する電流の限界値
        keithley.write(':SOURce:VOLTage:ILIMit ' + '{:E}'.format(float(measurement_condition_List[5])))
        # 測定の設定
        keithley.write(':SOURce:SWEep:VOLTage:LINear ' + str(measurement_condition_List[0]) + ', ' + str(measurement_condition_List[1]) + ', ' + str(measurement_condition_List[3]) + ', ' + '{:E}'.format(float(measurement_condition_List[2])))
        # FRONtでの測定に固定
        keithley.write(':ROUTe:TERMinals FRONt')
        # 測定を仕掛ける際の条件みたいなやつ
        keithley.write(':trig:bloc:bran:even 6, disp, 12')
        keithley.write('*CLS')
        keithley.write(':STATus:OPERation:MAP 0,2732,2731')
        keithley.write(':TRACe:CLEar "defbuffer1"')
        keithley.timeout = 240000
        keithley.write(':INITiate')
        keithley.write('*WAI')
        while True:
            out = keithley.query(':STATus:OPERation:CONDition?')
            out = out.rstrip('\n')
            if out == '0':
                break
        
        out = keithley.query(':OUTPut:STATe?')
        out = out.rstrip('\n')
        if out == '1':
            keithley.write(':ABORt')
            keithley.write(':OUTPut OFF')
        out = keithley.query(':TRACe:ACTual? "defbuffer1"')
        measure_points = int(out.rstrip('\n'))
        # ここに相対時間、電圧値、電流値が入る
        rawdata = []
        rawdata_str = []
        for i in range(1, measure_points + 1):
            sCmds = f':TRACe:DATA? {i}, {i}, "defbuffer1", RELative, SOURce, READing'
            out = keithley.query(sCmds)
            data_str = out.rstrip('\n').split(',')
            data = [float(n) for n in data_str]
            rawdata.append(data)
            rawdata_str.append(data_str)
        
        # 通信を閉じる
        keithley.close()
        rm.close()

        # ファイルへの保存
        with open( fr"{data_path}\IV_data\{measurement_number}.csv", mode = 'w') as f:
            f.write('測定開始日時,' + str(nowtime) + '\n')
            f.write('温度(℃),' + temp + ',pO2(atm),' + pO2 + '\n')
            f.write('経過時間[秒],電圧[V],電流[A],抵抗値[Ω]\n')
            for i in range(measure_points):
                f.write(f"{rawdata_str[i][0]},{rawdata_str[i][1]},{rawdata_str[i][2]},{rawdata[i][1] / rawdata[i][2]}\n")
        
        print(nowtime)
        print('Temp(C):' + temp)
        if pO2 != '':
            print('Log pO2:-' + Log_pO2)
        voltage = [row[1] for row in rawdata]
        current = [row[2] for row in rawdata]

        x = np.array(current)
        y = np.array(voltage)
        result = calculate_linear_regression(x, y)
        resist = result[0]
        s = ((sample_info[2] / 20) ** 2) * np.pi
        l = sample_info[3] / 10
        sigma = l / (s * resist)

        if sigma > 0:
            cond_result = np.log10(sigma)
        else : cond_result = "error"
        print(f"Log σ: {cond_result}")

        # エクセル更新用に今の測定状態を記録するファイル
        with open(fr"{data_path}\status_now.txt" , mode = 'w') as f:
            f.write(str(measurement_number) + '\n' + str(measure_points))

        print('待機中…ロガーファイルの切り替えや測定終了する場合は「Ctrl + C」を押してください。')
        print('Waiting...Press “Ctrl + C” to switch logger files or to exit measurement.')

    except Exception as e:  # 通信断などのエラーをキャッチ
        print(f'\n【エラー発生】測定中に通信エラーを検知しました。ケーブルやケースレーを確認してください: {e}')
        print(f'\n [Error] A connection error was detected during measurement. Please check the cable and the Keithley instrument: {e}')
        has_error = True  # エラーフラグを立てる
       
# 最小二乗法の傾きaと切片bを求めるやつ
def calculate_linear_regression(x, y):
    n = len(x)
    a = ((np.dot(x, y)- y.sum() * x.sum()/n)/
        ((x ** 2).sum() - x.sum()**2 / n))
    b = (y.sum() - a * x.sum())/n
    return [a, b]

# ===================================本体はここから=================================================

has_error = False  # エラーフラグの初期化

# 起動時処理、keithleyの自動検索と接続確認、最後にパスを渡す
keithley_path, furnace_number = initialize_measurement()
datapath = DATA_DIRECTORY_MAP[furnace_number]
print(f"この測定は{furnace_number}号機での測定です。")
print(f"This measurement was taken on Unit {furnace_number}.")

# 誰のディレクトリにデータを置くか選択
datapath = select_save_directory(datapath)
# サンプルネームを記入してもらいフォルダを作る
print('サンプル名を英数字のみで入力してください。')
print('Please enter sample name in alphanumeric characters only.')
while True:
    sample_name = input('sample name:')
    if sample_name.isalnum():
        break
    print('サンプル名は英数字のみで記入してください。')
    print('Sample names must be alphanumeric only.')

print('英数字のみでのコメントがある場合は入力してください。何もなければenterを押してください。')
print('Enter alphanumeric comments only, if any. If none, press enter.')
while True:
    sample_comment = input('comments:')
    if sample_comment.isalnum() or sample_comment == '':
        break
    print('コメントは英数字のみで記入してください。何もなければenterを押してください。')
    print('comments must be alphanumeric only. If none, press enter.')

datapath = create_save_folder(datapath, sample_name, sample_comment, furnace_number)

# サンプル情報を入力してもらう、重さ、高さ、直径、端子間距離
sample_info = input_sample_info()

with open(fr"{datapath}\sample_info.txt", mode = 'w') as f:
    f.write(f"{sample_name}\n")
    f.write(f"{sample_info[0]}\n")
    f.write(f"{sample_info[1]}\n")
    f.write(f"{sample_info[2]}\n")
    f.write(f"{sample_info[3]}")

measurement_condition_list = load_measurement_conditions(datapath)

# ロガーファイル等の情報、ロガー、チャンネル番号、6号機ロガー
logger_infoList = load_logger_info(datapath, furnace_number)

# 測定準備
os.mkdir(fr"{datapath}\IV_data")
measurement_number = 0
# 測定実行
while True:
    try:
        while True:
            if has_error:
                print('\n機器との再接続を試みています...')
                print('\nAttempting to reconnect to the device...')
                try:
                    rm = pyvisa.ResourceManager()
                    visa_list = rm.list_resources()
                    identifier = KEITHLEY_MAP[furnace_number]  # 最初に選んだ号機の識別子
                    new_path = None
                    for machine_name in visa_list:
                        if identifier in machine_name:
                            new_path = machine_name
                            break
                    rm.close()
                    
                    if new_path:
                        keithley_path = new_path
                        has_error = False  # 復活したらフラグを下ろす
                        print('再接続に成功しました。測定を再開します。')
                        print('Reconnection successful. Resuming measurement.')
                    else:
                        print('機器が見つかりません。次の測定周期に再試行します。')
                        print('Device not found. Retrying at the next measurement cycle.')
                except Exception as re_err:
                    print(f'再接続処理中にエラーが発生しました: {re_err}')
                    print(f'An error occurred during the reconnection process.: {re_err}')
            
            # エラーが出ていない場合のみ測定を実行する
            if not has_error:
                measurement_number += 1
                nowtime = datetime.datetime.now()
                thread1 = Thread(target = measure_conductivity, args = (nowtime, measurement_number, keithley_path, datapath, furnace_number, measurement_condition_list, logger_infoList, sample_info), daemon = True)
                thread1.start()
                time.sleep(measurement_condition_list[8])
                thread1.join()
            else:
                # 機器が見つからずエラーフラグが立ったままの場合は、測定をスキップして5分待機
                time.sleep(measurement_condition_list[8])


    except KeyboardInterrupt:
        print('測定を終了する場合は「q」を、酸素センサーの切り替えやロガーファイルを切り替える場合は「l」を押してください。')
        print('Press “q” to exit the measurement, or “l” to switch oxygen sensors or logger files.')
        while True:
            answer = input('q/l:')
            if answer == 'q':
                exit()
            elif answer == 'l':
                with open(fr"{datapath}\logger_files.txt", mode = 'r') as f:
                    old_data = f.read()
            
                with open(fr"{datapath}\logger_files_old{measurement_number}.txt", mode = 'w') as f:
                    f.write(old_data)

                logger_infoList = load_logger_info(datapath, furnace_number)
                break

            else:
                print('q/l で回答してください。')
                print('Please answer in q/l.')
