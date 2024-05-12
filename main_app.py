import numpy as np
import cv2
import os
import pandas as pd
import datetime
import re
import time
from progress.spinner import Spinner
from progress.bar import IncrementalBar
import xlrd
from xlutils.copy import copy
from itertools import takewhile
import calendar
import warnings
warnings.filterwarnings("ignore")
1



name_file_log='Данные отгрузки.xls'
dir_file_log='./log_data/'
name_file_conf='config_data_save.txt'
dir_file_conf='./config/'
dir_video='./video/'
ful_dir_exl=0

class bcolors:
    HEADER    = '\033[95m'
    OKBLUE    = '\033[94m'
    OKGREEN   = '\033[92m'
    WARNING   = '\033[93m'
    FAIL      = '\033[91m'
    BOLD      = '\033[1m'
    UNDERLINE = '\033[4m'
    ENDC      = '\033[0m'

    # Method that returns a message with the desired color
    # usage:
    #    print(bcolor.colored("My colored message", bcolor.OKBLUE))
    @staticmethod
    def colored(message, color):
      return color + message + bcolors.ENDC

    # Method that returns a yellow warning
    # usage:
    #   print(bcolors.warning("What you are about to do is potentially dangerous. Continue?"))
    @staticmethod
    def warning(message):
      return bcolors.WARNING + message + bcolors.ENDC

    # Method that returns a red fail
    # usage:
    #   print(bcolors.fail("What you did just failed massively. Bummer"))
    #   or:
    #   sys.exit(bcolors.fail("Not a valid date"))
    @staticmethod
    def fail(message):
      return bcolors.FAIL + message + bcolors.ENDC

    # Method that returns a green ok
    # usage:
    #   print(bcolors.ok("What you did just ok-ed massively. Yay!"))
    @staticmethod
    def ok(message):
      return bcolors.OKGREEN + message + bcolors.ENDC

    # Method that returns a blue ok
    # usage:
    #   print(bcolors.okblue("What you did just ok-ed into the blue. Wow!"))
    @staticmethod
    def okblue(message):
      return bcolors.OKBLUE + message + bcolors.ENDC

    # Method that returns a header in some purple-ish color
    # usage:
    #   print(bcolors.header("This is great"))
    @staticmethod
    def header(message):
      return bcolors.HEADER + message + bcolors.ENDC

def check_file(dir, name_file):
    list_file = os.listdir(path=dir)
    if name_file in list_file:
        return True
    else:
        return  False
def search_by_file_extension(file_last_date=1):
    list_file_xls=[]
    for file in os.listdir(dir_file_log):
        if file.endswith(".xls"):
            list_file_xls.append(file)
    if file_last_date==1:
        return  list_file_xls[-1]
    else:
        return list_file_xls

def index_last_simbol(string, symbol):
    counter_index=0
    check_count_symbol=0
    count_symbol = string.count(symbol)
    for i in string:

        if i==symbol and check_count_symbol<count_symbol:
            check_count_symbol+=1
        elif check_count_symbol==count_symbol:
            return counter_index
        counter_index+=1

def column_len(sheet, index):
    col_values = sheet.col_values(index)
    col_len = len(col_values)
    for _ in takewhile(lambda x: not x, reversed(col_values)):
        col_len -= 1
    return col_len

def read_config_data(path):
    flag_correct=0
    dir_excel=''
    with open(path) as f:
        data=f.read()
        try:
            year=re.search('Year=\d', data)[0]
            month = re.search('Month=\d', data)[0]
            dir_excel = re.split(month, data)[1]
            dir_excel=re.split(r'=',dir_excel)[1]
            year = int(re.search('\d', year)[0])
            month = int(re.search('\d', month)[0])
            id = index_last_simbol(dir_excel, '.')
            if dir_excel[(id - 6)]=='-':
                log_year = int(dir_excel[(id - 10):id - 6])
            else:
                log_year = int(dir_excel[(id - 5):id-1])
            new_date = datetime.date.today()
            div_year = new_date.year - log_year
            if div_year == year:
                write_config_data(step_month=month, step_year=year)
                year, month, dir_excel = read_config_data(path)

        #re.split(r'y', 'Analytics')
        except (TypeError,IndexError):

            name_file_log_ = '0'
            new_date = datetime.date.today()
            name_file_xls=search_by_file_extension()
            id = index_last_simbol(name_file_xls, '.')
            if name_file_xls[(id - 6)]=='-':
                step_year = int(name_file_xls[(id - 5):id-1])-int(name_file_xls[(id - 10):id - 6])+1
                div_year = (new_date.year - int(name_file_xls[(id - 5):id-1]))

            else:
                step_year = 1
                div_year = new_date.year - int(name_file_xls[(id - 5):id - 1])

            if div_year<=0:
                name_file_log_= name_file_xls
            year=step_year
            month=2
            write_config_data(step_month=month,step_year=year, name_file_log_=name_file_log_)
            flag_correct=1
            text='FILE CONFIG CORECTING:    '
            progress(text)


    if flag_correct>0:
        year, month, dir_excel=read_config_data(path)


    return year, month, dir_excel

def write_config_data(step_month, step_year, name_file_log_='0'):
    now_date = datetime.date.today()
    id_ext = name_file_log.index('.', 2)  # заменить name_file_log_
    if step_year<2 and name_file_log_=='0':
        name_file_log_ = name_file_log[:id_ext] + str(now_date.year) + name_file_log[id_ext:]
    elif step_year>=2and name_file_log_=='0':
        name_file_log_ = name_file_log[:id_ext] + str(now_date.year)+'-'+ str(now_date.year+step_year) + name_file_log[id_ext:]
    strings_config = {'Year': str(step_year), 'Month': str(step_month), 'Path': dir_file_log + name_file_log_}
    with open(dir_file_conf + name_file_conf, 'w') as file:
        for i in strings_config:
            if i=='Path':
                file.write(i + '=' + strings_config[i])
            else:
                file.write(i + '=' + strings_config[i] + '\n')

def open_log(dir):
    log_data = xlrd.open_workbook(dir, on_demand=True, formatting_info=True)
    id_sheet = len(log_data.sheet_names())
    return  id_sheet, log_data

def creat_excel(dir, month):
    name_index_list = ['id', 'Дата', 'Время', 'Количество особей']
    log_data = pd.DataFrame(columns=name_index_list)
    now_date = datetime.date.today()
    if month<2:
        name_sheet = calendar.month_abbr[now_date.month]
    else:
        name_sheet = calendar.month_abbr[now_date.month] + '-' + calendar.month_abbr[now_date.month + month]
    log_data.to_excel(dir, sheet_name=name_sheet, index=0)

def sorting_name(list_file):
    temp=0
    for index in range(len(list_file)):
        index_point_cur=index_last_simbol(list_file[index],'.')
        num_curent = int(list_file[index][index_point_cur - 2])
        for index_next in range(index+1, len(list_file)):
            index_point_next=index_last_simbol(list_file[index_next],'.')
            num_next = int(list_file[index_next][index_point_next - 2])
            if num_next<num_curent:
                temp=list_file[index]
                list_file[index]=list_file[index_next]
                list_file[index_next]=temp
    return  list_file

def max_index(list_file, indent):
    max_index_file=0
    index_list_of_mif=0
    for index_list, name_file in enumerate(list_file):

        if name_file.find(indent) > -1:
            index_point = index_last_simbol(name_file, '.')
            index_file = int(list_file[index_list][index_point - 2])
            if max_index_file<index_file:
                max_index_file=index_file
                index_list_of_mif=index_list

    return max_index_file, index_list_of_mif




def prep_video_name():
    all_list_file = os.listdir(path='./video')
    video_list_file=[x for x in all_list_file if (x.find('.mp4')>-1 or x.find('.avi')>-1)]
    if len(video_list_file)==0:
        return 0
    size_video_list_file=len(video_list_file)
    last_video_file=1
    flag_new_video=0
    for i in video_list_file:
        if last_video_file-1==size_video_list_file:
            video_list_file = sorting_name(video_list_file)
            return video_list_file
        elif i.find('tv_')>-1 :
            last_video_file+=1
    max_index_video, index_in_vlf=max_index(video_list_file, 'tv_')
    if max_index_video>=last_video_file:
        index = index_last_simbol(video_list_file[index_in_vlf], '.')
        app_ = video_list_file[index_in_vlf][index - 1:]
        name = video_list_file[index_in_vlf][:index - 6]
        new_name = name + '_tv_' + str(last_video_file-1) + app_
        os.rename('./video/' + str(video_list_file[index_in_vlf]), './video/' + new_name)



    print("Добавлен(ы) видео файл(ы):")
    for i in video_list_file:
        if i.find('tv_')==-1:
            index=index_last_simbol(i,'.')
            app_=i[index-1:]
            name=i[:index-1]
            new_name=name+'_tv_'+str(last_video_file)+app_
            print(new_name)
            os.rename('./video/'+str(i), './video/'+new_name)
            last_video_file+=1
    all_list_file = os.listdir(path='./video')
    video_list_file = [x for x in all_list_file if (x.find('.mp4') > -1 or x.find('.avi') > -1)]
    video_list_file=sorting_name(video_list_file)
    return video_list_file


def estimate_date_log(dir, month):
    id_sheet, log_data=open_log(dir)
    sheet = log_data.sheet_by_index(id_sheet - 1)
    id = column_len(sheet, 1)
    log_data2 = copy(log_data)
    if id==1:
        id_sheet = len(log_data.sheet_names())
        log = log_data2.get_sheet(id_sheet - 1)
        log_data.release_resources()
        del log_data
        return log, log_data2, id
    else:
        #######считать года################
        #last_date = xlrd.xldate.xldate_as_datetime(sheet.row_values(id-1)[0],log_data.datemode)
        last_date=sheet.cell(id-1,1).value
        new_date = datetime.date.today()
        last_date = datetime.datetime.strptime(last_date, '%Y.%m.%d')
        div_month = new_date.month - last_date.month
        if div_month<0:
            div_month=12+div_month
        if div_month == month:
            name_index_list = ['id', 'Дата', 'Время', 'Количество особей']
            if month < 2:
                name_sheet = calendar.month_abbr[new_date.month]
            else:
                name_sheet = calendar.month_abbr[new_date.month] + '-' + calendar.month_abbr[new_date.month + month]
            log_data2.add_sheet(str(name_sheet))
            log_data.release_resources()
            del log_data
            log_data2.save(dir)
            print("Creat sheet successfully")
            id_sheet, log_data = open_log(dir)
            id_sheet = len(log_data.sheet_names())
            log_data2 = copy(log_data)
            log = log_data2.get_sheet(id_sheet - 1)
            log.write(0, 0, name_index_list[0])
            log.write(0, 1, name_index_list[1])
            log.write(0, 2, name_index_list[2])
            log.write(0, 3, name_index_list[3])
            log_data.release_resources()
            del log_data
            return log, log_data2, 1
        else :
            id_sheet = len(log_data.sheet_names())
            log = log_data2.get_sheet(id_sheet - 1)
            log_data.release_resources()
            del log_data
            return  log, log_data2,id

def log_save(cout_fish):
    path_config_file=dir_file_conf+name_file_conf
    year, month,dir_and_name_file=read_config_data(path_config_file)
    index = index_last_simbol(dir_and_name_file, '/')
    dir = dir_and_name_file[:index]
    name_file = dir_and_name_file[index:]
    list_file = os.listdir(path=dir)
    if name_file not  in list_file:
        creat_excel(dir_and_name_file, month)
        """ ДОПИСАТЬ СОХРАНЕНИЕ В ФАЙЛ ПУТИ"""
    log, log_data2, id=estimate_date_log(dir_and_name_file, month)
    log.write(id, 0, id)
    log.write(id, 1, str(datetime.date.today().strftime("%Y.%m.%d")))
    time_now = datetime.datetime.now()
    log.write(id, 2, str(datetime.time(int(time_now.hour), int(time_now.minute), int(time_now.second))))
    log.write(id, 3, cout_fish)
    # log_data.release_resources()
    # del log_data
    # if (os.path.exists(dir_and_name_file)):
    #     os.remove(dir_and_name_file)
    log_data2.save(dir_and_name_file)
    print("Save log successfully")

def camera_view():
    cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)
    while(1):
        fl=input('ОТРАЗИТЬ ИЗОБРАЖЕНИЕ? д/н ')
        if fl=='д':
            break
        elif fl=='н':
            break
        else:
            continue
    while (cap.isOpened()):
        ret, frame = cap.read()
        if fl == 'д':
            frame=cv2.flip(frame,1)
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        imS = cv2.resize(gray, (960, 540))
        cv2.imshow('web_camera', imS)
        k = cv2.waitKey(33)
        if k == 27:  # Esc key to stop
            cap.release()
            cv2.destroyAllWindows()
        time.sleep(0.001)


def start_count(mode=1, name_video='0'):

    import matplotlib.pyplot as mpl
    fig, ax=mpl.subplots()
    EntranceCounter = 0
    MinCountourArea = 2500
    MaxCountourArea = 67000
    QttyOfContours = 0
    Cutoff=500
    Cutstart=122
    size_x=960
    size_y=540
    h_cut = 0
    w_cut = 0
    y_cut = 0
    flag_cut = 0
    new_fish = 0
    kernel = np.ones((8, 8), np.uint8)
    kernel2 = np.ones((7, 7), np.uint8)#8,8
    CoordXCentroid_list_past = []
    CoordXCentroid_list_new = []
    CoordXCentroid_list_div = []
    while (1):
        fl = input('ОТРАЗИТЬ ИЗОБРАЖЕНИЕ? д/н ')
        if fl == 'д':
            break
        elif fl == 'н':
            break
        else:
            continue
    bar = IncrementalBar("  Подготовка видео  ", max=5)
    event_start = 0
    if mode==1 and name_video=='0':
        cap = cv2.VideoCapture(0,cv2.CAP_DSHOW)

    if mode==2 and name_video!='0':
        cap = cv2.VideoCapture(dir_video+name_video)
    try:
        cadr=0
        while(cap.isOpened()):
            ret, frame = cap.read()
            if fl == 'д':
                frame = cv2.flip(frame, 1)
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            imS = cv2.resize(gray, (size_x, size_y))


            while flag_cut<5:


                bar.next()
                ret, frame = cap.read()
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                imS = cv2.resize(gray, (size_x, size_y))#960, 540
                ret_cut, thresh_cut=cv2.threshold(imS, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
                height_cut = np.size(thresh_cut, 0)
                width_cut = np.size(thresh_cut, 1)
                thresh_cut = cv2.dilate(thresh_cut, None, iterations=2)
                cnt_cut, _ = cv2.findContours(thresh_cut.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                for c in cnt_cut:
                    #print( cv2.contourArea(c))
                    if cv2.contourArea(c) > MaxCountourArea:
                        (x, y, w_cut, h_cut) = cv2.boundingRect(c)
                        if y>20:
                            y_cut=y
                flag_cut+=1
                #time.sleep(0.04)

            if cadr<0:
                cadr+=1
            else:
                cadr=0
                ret, thresh = cv2.threshold(imS, 50, 140, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

                height = np.size(thresh, 0)
                width = np.size(thresh, 1)
                # plot reference lines (entrance and exit lines)

                thresh = cv2.dilate(thresh, None, iterations=2)
                thresh[y_cut+30:h_cut + y_cut, 0:w_cut] = 0
                thresh[0:h_cut-30, 0:w_cut] = 0#-190
                thresh=cv2.erode(thresh, kernel, iterations=2)
                thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel2)

                color_thresh = cv2.cvtColor(thresh, cv2.COLOR_GRAY2RGB)
                cnts, _ = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                for c in cnts:
                    # if a contour has small area, it'll be ignored
                   # print(cv2.contourArea(c))
                    if cv2.contourArea(c) < MinCountourArea :

                        continue
                    QttyOfContours = QttyOfContours + 1
                    ###########################################
                    o=-1
                    if  EntranceCounter== o:
                       print(cv2.contourArea(c))
                       rect0 = cv2.minAreaRect(c)
                       box0 = cv2.boxPoints(rect0)
                       box0 = np.int0(box0)
                       color_thresh0 = cv2.drawContours(color_thresh, [box0], 0, (0, 255, 0), 2)
                       color_thresh0 = cv2.resize(color_thresh0, (0, 0), fx=0.6, fy=0.6)

                       cv2.imshow('frame0', color_thresh0)
                    # if EntranceCounter==500:
                    #     cv2.imwrite('D:/Work/fish/result_video/tv_2_92_f.png',imS)
                    ##########################################################################
                    rect = cv2.minAreaRect(c)
                    box = cv2.boxPoints(rect)
                    box = np.int0(box)
                    color_thresh = cv2.drawContours(color_thresh, [box], 0, (0, 0, 255), 2)
                    Xs = [i[0] for i in box]
                    Ys = [i[1] for i in box]
                    x = min(Xs)
                    w = max(Xs)
                    y = min(Ys)
                    h = max(Ys)
                    # center = (int((x1 + x2) / 2), int((y1 + y2) / 2))
                    # cv2.circle(color_thresh, center, 10, (0, 255, 0), -1)  # again this was mostly for debugging purposes
                    ###########################################
                    # draw an rectangle "around" the object
                    (x_text, y_text, w_text, h_text) = cv2.boundingRect(c)
                   # cv2.rectangle(color_thresh, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.putText(color_thresh, "w=: {}".format(str(int(rect[1][0]))),
                                (x_text, y), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)
                    cv2.putText(color_thresh, "h=: {}".format(str(int(rect[1][1]))),
                                (x_text, y+h_text), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)

                    # find object's centroid
                    # CoordXCentroid = (x + x + w) / 1.7
                    # CoordYCentroid = (y + y + h) / 2
                    CoordXCentroid=(x + w)/2
                    CoordYCentroid = (y + h)/2
                    if rect[1][1]>rect[1][0]:
                        relation_w_h=rect[1][0]/rect[1][1]
                    elif rect[1][1]<rect[1][0]:
                        relation_w_h = rect[1][1] / rect[1][0]
                    cv2.putText(color_thresh, "w/h=: {}".format(str('%.2f'% relation_w_h)),
                                (int(CoordXCentroid), int(CoordYCentroid)), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)
                    if len(CoordXCentroid_list_past)==0:
                        CoordXCentroid_list_past.append(CoordXCentroid)
                    elif CoordXCentroid<Cutoff and (CoordXCentroid>Cutstart or (CoordXCentroid-min(CoordXCentroid_list_past)>0\
                                                                            or CoordXCentroid-min(CoordXCentroid_list_past)<-30 )):#762:
                        CoordXCentroid_list_new.append(CoordXCentroid)
                    ObjectCentroid = (int(CoordXCentroid), int(CoordYCentroid))
                    cv2.circle(color_thresh, ObjectCentroid, 1, (0, 0, 255), 5)


                if len(CoordXCentroid_list_new)>0 and \
                        (len(CoordXCentroid_list_new)==len(CoordXCentroid_list_past) or len(CoordXCentroid_list_new)>len(CoordXCentroid_list_past)):
                    div_len=len(CoordXCentroid_list_new)-len(CoordXCentroid_list_past)
                    CoordXCentroid_list_new.sort()
                    for i in  range(div_len):
                        CoordXCentroid_list_past.append(CoordXCentroid_list_new[i])
                        CoordXCentroid_list_new[i]=0
                    CoordXCentroid_list_past.sort()
                    CoordXCentroid_list_div = [CoordXCentroid_list_new[n] - CoordXCentroid_list_past[n] for n in
                                               range(0, len(CoordXCentroid_list_new))]
                    if div_len==0:
                        CoordXCentroid_list_past=CoordXCentroid_list_new.copy()


                    new_fish=sum(i < 0 for i in CoordXCentroid_list_div)
                    if new_fish>0:
                        EntranceCounter=EntranceCounter+new_fish
                        new_fish=0


                elif event_start==0 and len(CoordXCentroid_list_past)>0:
                     EntranceCounter=len(CoordXCentroid_list_past)
                     event_start=1
                elif len(CoordXCentroid_list_new)<len(CoordXCentroid_list_past) and len(CoordXCentroid_list_new)>0 :

                    ##########################Пробная вставка#########################
                    CoordXCentroid_l_p_min=min(CoordXCentroid_list_past)
                    for new_x in CoordXCentroid_list_new:
                        if new_x<CoordXCentroid_l_p_min:
                            EntranceCounter = EntranceCounter + 1

                    ##################################################################
                    CoordXCentroid_list_past = CoordXCentroid_list_new.copy()
                #CoordXCentroid_list_past = [i for i in CoordXCentroid_list_past if i <= 863]
                list.clear(CoordXCentroid_list_new)
                list.clear(CoordXCentroid_list_div)
                imS = cv2.cvtColor(imS, cv2.COLOR_GRAY2BGR)
                height = np.size(color_thresh, 0)
                width = np.size(color_thresh, 1)
                cv2.line(color_thresh, (0, height), (width, height), (0, 255, 0), 2)
                cv2.line(color_thresh, (Cutoff, height), (Cutoff, 0), (0, 255, 0), 2)
                color_thresh=np.concatenate((color_thresh, imS), axis=0)
                cv2.putText(color_thresh, "Count: {}".format(str(EntranceCounter)),
                            (10, 50),cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)
                color_thresh = cv2.resize(color_thresh, (0, 0), fx=0.6, fy=0.6)
                #cv2.destroyAllWindows()
                cv2.imshow('frame',color_thresh)
                #time.sleep(0.1)
                k = cv2.waitKey(33)
                if k == 27:  # Esc key to stop
                    print("\ncount_fish=", EntranceCounter)
                    cap.release()
                    cv2.destroyAllWindows()
                    return EntranceCounter
                # if cv2.waitKey(1) & 0xFF == 27#ord(27):
                #     print("count_fish=", EntranceCounter)
                #     cap.release()
                #     cv2.destroyAllWindows()
                #     return EntranceCounter
    except cv2.error:
        cap.release()
        cv2.destroyAllWindows()
        print("count_fish=", EntranceCounter)
        return EntranceCounter



def menu():
    print("СПИСОК КОМАНД:")
    list_command=[' 1-запуск подсчета',
                  ' 2-проверка на видео',
                  ' 3-вид с камеры ',
                  ' 0-Закрыть программу',
                  ' Esc-выключить видео/камеру',
                  ]
    for i in range(len(list_command)):
        print("<<<<<< "+list_command[i])

def preparation ():
    ###CHECK CONF FILE######
    flag_prep_video_name=1
    if not  os.path.exists(dir_file_conf):
        os.mkdir('config')
        text = "CREATING PATH 'config   '"
        progress(text)
        print("PATH 'config' CREATED SUCCESSFULLY")
    if not  os.path.exists('./log_data/'):
        text = "CREATING PATH 'log_data'"
        progress(text)
        os.mkdir('log_data')
        print("PATH 'log_data' CREATED SUCCESSFULLY")

    if not os.path.exists('./video/'):
        text = "CREATING PATH 'video'"
        progress(text)
        os.mkdir('video')
        print("PATH 'video' CREATED SUCCESSFULLY")
        flag_prep_video_name=0

    if check_file(dir_file_conf, name_file_conf) ==False:
        year = 1
        month = 2
        write_config_data(step_month=month,step_year=year)
        year, month, dir_file_log = read_config_data(dir_file_conf + name_file_conf)
        print("CREAT  file 'config.txt'")

    elif check_file(dir_file_conf, name_file_conf)==True:
        year, month, dir_file_log = read_config_data(dir_file_conf+name_file_conf)







    print(f"    Шаг месяца={month}, Шаг года={year}")
    ###########################################
    ###CHECK LOG FILE######
    index=index_last_simbol(dir_file_log, '/')
    dir = dir_file_log[:index]
    name=dir_file_log[index:]
    if check_file(dir=dir, name_file=name)==False:
        creat_excel(dir=dir_file_log, month=month)
        print("CREAT  log file")

def main():

    preparation()
    while(1):
        menu()
        command = input('Введите команду    ')
        if int(command) == 1:
            cout_fish=start_count(1)
            log_save(cout_fish)
        elif int(command)==2:
            video_list = prep_video_name()
            if len(video_list)>0:
                print("\n*******СПИСОК ВИДЕО ФАЙЛОВ*******")
                for i in video_list:
                    print(i)
                while(1):
                    print("Выйти из режима видео  0")
                    test_video = int(input('Запустить видео с индексом tv_ '))
                    if test_video==0:
                        break
                    try:
                        test_video=video_list[test_video-1]
                        while(1):
                            flag_save = input('Сохранить результаты подсчета видео? д/н ')
                            if flag_save == 'д':
                                break
                            elif flag_save== 'н':
                                break
                            else:
                                continue
                        cout_fish = start_count(2, test_video)
                        if flag_save=='д':
                            log_save(cout_fish)
                        print(f"количество рыб на {test_video}={cout_fish}")
                    except IndexError:
                        print("***************")
                        print(f"Видео с индексом tv_{test_video} не существует")
                        print("***************")
            elif video_list==False:
                print("В папке video нет файлов .avi или .mp4")
        elif int(command)==3:
            camera_view()
        elif int(command)==0:
            break
        else :
            print("введена несуществующая команда")
            pass


def progress(text):
    spinner = Spinner(text)
    for item in range(20):
        spinner.next()
        time.sleep(0.04)

    spinner.finish()

if __name__ == '__main__':
    main()

