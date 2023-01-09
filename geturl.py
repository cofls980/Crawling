import tkinter as tk
import tkinter.ttk as ttk
import traceback
from dataclasses import dataclass
from tkinter import filedialog
import tkinter.messagebox as msgbox
import pandas as pd
from bs4 import BeautifulSoup
from urllib.request import urlopen
import threading

import json
import time
from tqdm import tqdm

import os.path

from datetime import datetime


# ####################################### utils ###########################################
def valid_file_combo():
    file_path = ui.box_find_file.get()
    if not os.path.exists(ui.data.default_live):
        print_in_list_box('필수_실시간.xlsx 파일 위치를 다시 확인해 주세요.')
        return ''
    if not os.path.exists(ui.data.default_view_like):
        print_in_list_box('필수_조회수+좋아요.xlsx 파일 위치를 다시 확인해 주세요.')
        return ''
    if file_path == '' or ui.combo_select_loop_time.get() == '반복 시간 선택':
        print_in_list_box('파일과 반복 시간 모두 선택해 주세요.')
        return ''
    if '.xlsx' not in file_path or '.xls' not in file_path:
        print_in_list_box('엑셀 파일을 선택해 주세요.')
        return ''
    return file_path


def valid_channel_list_excel_form(file_path):
    rd_df = pd.read_excel(file_path)
    try:
        names = list(rd_df['채널 이름'])
        urls = list(rd_df['실시간 주소'])
        if len(names) != len(urls):
            return [], []
        for i in range(len(names)):
            if type(names[i]) == float and type(urls[i]) != float:
                print_in_list_box('엑셀 파일 형식을 확인해 주세요.')
                return [], []
            if type(names[i]) != float and type(urls[i]) == float:
                print_in_list_box('엑셀 파일 형식을 확인해 주세요.')
                return [], []
    except:
        print_in_list_box('엑셀 파일 형식을 확인해 주세요.')
        return [], []
    return names, urls


def valid_channel_url_form(name, url):
    try:
        if not url.startswith('https://www.youtube.com/') or not url.endswith('/streams'):
            print_in_list_box(str(name) + ': URL 형식 확인해 주세요.')
            return False
    except:
        print_in_list_box('엑셀 파일 형식을 확인해 주세요.')
        return False
    return True


def get_curr_time():
    now = str(datetime.now().hour) + ':' + str(datetime.now().minute) + ':' + str(datetime.now().second)
    return now


def print_in_list_box(comment):
    ui.box_result_comment.insert(tk.END, '[' + ui.data.today + ' ' + get_curr_time() + '] ' + comment)


def open_live_excel(data):
    if data.live_index != 0:
        df_live = pd.read_excel(ui.data.real_live_name + str(data.live_index) + '.xlsx')
    else:
        df_live = pd.read_excel(ui.data.default_live)
    return df_live


def find_video_row_index(ch_name, video_id, data):
    for i in range(len(data)):
        if data.loc[i]['채널명'] == ch_name and data.loc[i]['영상 주소'] == video_id:
            return i
    return -1


# ####################################### utils ###########################################

# ####################################### program ###########################################
def crawling(name, url):
    try:
        html = urlopen(url)
        result = html.read()
        soup = BeautifulSoup(result, 'html.parser')
        soup = soup.body
        soup = str(soup.prettify())
        lines = soup.split('\n')
        del soup
    except:
        print_in_list_box(str(name) + ': 해당 채널이 존재하지 않습니다.')
        return
    return lines


def live_scraping(parsing, data, channel_name, column_name):
    for line in parsing:
        if 'ytInitialData' in line:
            line = line.strip('var ytInitialData = ')
            line = line.strip(';')
            json_object = json.loads(line)
            tabs = json_object.get('contents').get('twoColumnBrowseResultsRenderer').get('tabs')
            for a_tab in tabs:
                if a_tab.get('tabRenderer') is None:
                    continue
                tab_render = a_tab.get('tabRenderer')
                if tab_render.get('title') == '실시간':
                    if tab_render.get('content').get('richGridRenderer') is None:
                        continue
                    contents = tab_render.get('content').get('richGridRenderer').get('contents')
                    for content in contents:  # 실시간 페이지에서 영상들 확인
                        if content.get('richItemRenderer') is None:
                            continue
                        if content.get('richItemRenderer').get('content') is None:
                            continue
                        if content.get('richItemRenderer').get('content').get('videoRenderer') is None:
                            continue
                        if content.get('richItemRenderer').get('content').get('videoRenderer').get(
                                'viewCountText') is None:
                            continue
                        if content.get('richItemRenderer').get('content').get('videoRenderer').get('viewCountText').get(
                                'runs') is None:
                            continue
                        video = content.get('richItemRenderer').get('content').get('videoRenderer')
                        if '시청 중' in str(video.get('viewCountText')):
                            href = str(video.get('navigationEndpoint').get('commandMetadata').
                                       get('webCommandMetadata').get('url'))
                            href = 'https://youtube.com' + href
                            video_title = (video.get('title').get('runs'))[0].get('text')
                            viewers = (video.get('viewCountText').get('runs'))[0].get('text')
                            if '명' in viewers:
                                viewers = str(viewers)[:str(viewers).find('명')]
                            excel_row_idx = find_video_row_index(channel_name, href, data.df_live)
                            if excel_row_idx != -1:
                                data.df_live.loc[excel_row_idx, '영상 제목'] = video_title
                                data.df_live.loc[excel_row_idx, column_name] = int(viewers.replace(',', ''))
                            else:
                                insert_values = []
                                for column in range(len(data.df_live.columns)):
                                    if column == 0:
                                        insert_values.append(channel_name)
                                    elif column == 1:
                                        insert_values.append(video_title)
                                    elif column == 2:
                                        insert_values.append(href)
                                    elif column == len(data.df_live.columns) - 1:
                                        insert_values.append(int(viewers.replace(',', '')))
                                    else:
                                        insert_values.append(int(0))
                                data.df_live.loc[len(data.df_live)] = insert_values
                    break
            break


def run():
    try:
        if ui.data.stop:
            return
        start = time.time()
        ui.data.df_live = open_live_excel(ui.data)
        print("====================== START ======================")
        curr_time = get_curr_time()
        ui.data.df_live[curr_time] = 0
        for x in range(len(ui.data.channel_list_urls)):
            if ui.data.stop:
                return
            ui.progress_var.set(100 * ((x + 1) / len(ui.data.channel_list_urls)))
            ui.progress_bar.update()
            if type(ui.data.channel_list_names[x]) == float:
                continue
            if not valid_channel_url_form(ui.data.channel_list_names[x], ui.data.channel_list_urls[x]):
                continue
            parsing = crawling(ui.data.channel_list_names[x], ui.data.channel_list_urls[x])
            if parsing is None:
                continue
            live_scraping(parsing, ui.data, ui.data.channel_list_names[x], curr_time)
        ui.data.live_index = ui.data.live_index + 1
        ui.data.df_live.to_excel(ui.data.real_live_name + str(ui.data.live_index) + '.xlsx', index=False)
        print_in_list_box(ui.data.real_live_name + str(ui.data.live_index) + '.xlsx')
        sub_time = time.time() - start
        print('총 실행 시간: %s 초' % sub_time)
        print("======================= END =======================")

        ui.data.time_sum = ui.data.time_sum + sub_time
        ui.data.time_cnt = ui.data.time_cnt + 1

        goal_time = 60 * int(ui.data.time_term)
        timer = ui.data.time_sum / ui.data.time_cnt
        timer = goal_time - timer
        if timer < 0:
            timer = 0
        ui.data.run_thread = threading.Timer(timer, run)
        ui.data.run_thread.daemon = True

        if ui.data.stop:
            return

        ui.data.run_thread.start()
    except:
        msgbox.showerror("에러", "에러 문의 주세요.")
        err = traceback.format_exc()
        ErrorLog(str(err))
        stop_program()


# ####################################### program ###########################################

# #################################  monitoring program #####################################
def is_not_found(contents):
    if '이 동영상을 더 이상 재생할 수 없습니다.' in str(contents):
        return True
    if '동영상을 재생할 수 없음' in str(contents):
        return True
    if '업로더가 삭제한 동영상입니다.' in str(contents):
        return True
    return False


def view_like_scraping(video_info, parsing):
    like_cnt = -1
    view_cnt = -1
    for line in parsing:
        if 'ytInitialData' in line:
            line = line.strip('var ytInitialData = ').strip(';')
            json_object = json.loads(line)
            if json_object.get('contents') is None:
                continue
            if json_object.get('contents').get('twoColumnWatchNextResults') is None:
                continue
            if json_object.get('contents').get('twoColumnWatchNextResults').get('results') is None:
                continue
            if json_object.get('contents').get('twoColumnWatchNextResults').get('results').get('results') is None:
                continue
            if json_object.get('contents').get('twoColumnWatchNextResults').get('results').get('results').get(
                    'contents') is None:
                continue
            contents = json_object.get('contents').get('twoColumnWatchNextResults').get('results').get('results').get(
                'contents')
            if is_not_found(contents):
                # f = open('test.txt', 'a', newline='', encoding='utf-16')
                # f.write(video_info[0] + '\n\n')
                # f.write(str(contents) + '\n\n\n\n')
                # f.close()
                like_cnt = '삭제된 영상'
                view_cnt = '삭제된 영상'
                break
            for content in contents:
                if '시청 중' in str(content):
                    return False
                if '좋아요' in str(content):
                    if content.get('videoPrimaryInfoRenderer') is None:
                        continue
                    content = content.get('videoPrimaryInfoRenderer')
                    if content.get('viewCount') is None:
                        continue
                    if content.get('viewCount').get('videoViewCountRenderer') is None:
                        continue
                    if content.get('viewCount').get('videoViewCountRenderer').get('viewCount') is None:
                        continue
                    views = content.get('viewCount').get('videoViewCountRenderer').get('viewCount')
                    if views.get('simpleText') is None:
                        continue
                    if '조회수' not in str(views):
                        break
                    views = views.get('simpleText')
                    view_cnt = str(views)[4:]
                    view_cnt = int(view_cnt[:view_cnt.find('회')].replace(',', ''))
                    if content.get('videoActions') is None:
                        continue
                    if content.get('videoActions').get('menuRenderer') is None:
                        continue
                    if content.get('videoActions').get('menuRenderer').get('topLevelButtons') is None:
                        continue
                    items = content.get('videoActions').get('menuRenderer').get('topLevelButtons')
                    for item in items:
                        if '좋아요' in str(item):
                            if item.get('segmentedLikeDislikeButtonRenderer') is None:
                                continue
                            like1 = item.get('segmentedLikeDislikeButtonRenderer')
                            if like1.get('likeButton') is None:
                                continue
                            if like1.get('likeButton').get('toggleButtonRenderer') is None:
                                continue
                            like2 = like1.get('likeButton').get('toggleButtonRenderer')
                            if like2.get('defaultText') is None:
                                continue
                            if like2.get('defaultText').get('simpleText') is not None:
                                if str(like2.get('defaultText').get('simpleText')) == '좋아요':
                                    like_cnt = '정보없음'
                                    break
                            if like2.get('defaultText').get('accessibility') is None:
                                continue
                            like3 = like2.get('defaultText').get('accessibility')
                            if like3.get('accessibilityData') is None:
                                continue
                            if like3.get('accessibilityData').get('label') is None:
                                continue
                            item = like3.get('accessibilityData').get('label')
                            like_cnt = int(str(item)[4:str(item).find('개')].replace(',', ''))
                        break
                    break
            break
    if view_cnt == -1 or like_cnt == -1:
        return False
    row_index = len(ui.data.df_view_like)
    ui.data.df_view_like.loc[row_index, '채널명'] = video_info[0]
    ui.data.df_view_like.loc[row_index, '영상 제목'] = video_info[1]
    ui.data.df_view_like.loc[row_index, '영상 주소'] = video_info[2]
    ui.data.df_view_like.loc[row_index, '조회수'] = view_cnt
    ui.data.df_view_like.loc[row_index, '좋아요 수'] = like_cnt
    return True


def monitoring():
    try:
        if ui.data.stop:
            return
        while ui.data.live_index == 0:
            continue
        start = time.time()
        backup_df_live = ui.data.df_live
        if ui.data.view_like_index != 0:
            ui.data.df_view_like = pd.read_excel(ui.data.real_view_like_name + str(ui.data.view_like_index) + '.xlsx')
        else:
            ui.data.df_view_like = pd.read_excel(ui.data.default_view_like)
        # 조회수+좋아요 엑셀에 이미 저장된 채널을 제외한 채널을 리스트에 삽입
        ui.data.check_video_list = {}
        for i in range(len(backup_df_live)):
            if ui.data.view_like_index == 0:
                ui.data.check_video_list[(backup_df_live.loc[i]['채널명'], backup_df_live.loc[i]['영상 제목'],
                                          backup_df_live.loc[i]['영상 주소'])] = False
                continue
            if find_video_row_index(backup_df_live.loc[i]['채널명'], backup_df_live.loc[i]['영상 주소'],
                                    ui.data.df_view_like) == -1:
                ui.data.check_video_list[(backup_df_live.loc[i]['채널명'], backup_df_live.loc[i]['영상 제목'],
                                          backup_df_live.loc[i]['영상 주소'])] = False
        # 크롤링하여 리스트에 저장된 채널이 종료되었는지 확인
        flag = False
        for key, value in ui.data.check_video_list.items():
            if ui.data.stop:
                return
            parsing = crawling(key[0] + '(' + key[1] + ')', key[2])
            if parsing is None:
                print_in_list_box(key[0] + '(' + key[1] + ') - 알 수 없는 페이지')
                continue
            # 해당 채널이 실시간인지 아닌지 확인 후 데이터 저장
            res = view_like_scraping(key, parsing)
            if not flag:
                flag = res
        if flag:
            ui.data.view_like_index = ui.data.view_like_index + 1
            ui.data.df_view_like.to_excel(ui.data.real_view_like_name + str(ui.data.view_like_index) + '.xlsx', index=False)
            print_in_list_box(ui.data.real_view_like_name + str(ui.data.view_like_index) + '.xlsx')
        sub_time = time.time() - start
        print('모니터링 총 실행 시간: %s 초' % sub_time)

        ui.data.monitoring_thread = threading.Timer(0, monitoring)
        ui.data.monitoring_thread.daemon = True

        if ui.data.stop:
            return

        ui.data.monitoring_thread.start()
    except:
        msgbox.showerror("에러", "에러 문의 주세요.")
        err = traceback.format_exc()
        ErrorLog(str(err))
        stop_program()


# #################################  monitoring program #####################################
def monday_view():
    start = time.time()
    if ui.data.view_like_index != 0:
        backup_df = pd.read_excel(ui.data.real_view_like_name + str(ui.data.view_like_index) + '.xlsx')
        ui.data.df_view_like = pd.read_excel(ui.data.default_view_like)
    else:
        print_in_list_box('조회수+좋아요 폴더 또는 조회수+좋아요 엑셀 파일의 위치를 확인해 주세요.')
        return
    ui.data.check_video_list = {}
    for i in range(len(backup_df)):
        ui.data.check_video_list[(backup_df.loc[i]['채널명'], backup_df.loc[i]['영상 제목'],
                                  backup_df.loc[i]['영상 주소'])] = False
    flag = False
    idx = 0
    for key, value in ui.data.check_video_list.items():
        ui.progress_var.set(100 * ((idx + 1) / len(ui.data.check_video_list)))
        ui.progress_bar.update()
        idx = idx + 1
        parsing = crawling(key[0] + '(' + key[1] + ')', key[2])
        if parsing is None:
            print_in_list_box(key[0] + '(' + key[1] + ') - 알 수 없는 페이지')
            continue
        # 해당 채널이 실시간인지 아닌지 확인 후 데이터 저장
        res = view_like_scraping(key, parsing)
        if not flag:
            flag = res
    if flag:
        ui.data.df_view_like.to_excel(ui.data.real_monday_name + '.xlsx', index=False)
        print_in_list_box(ui.data.real_monday_name + '.xlsx')
    sub_time = time.time() - start
    print('월요일 조회수 총 실행 시간: %s 초' % sub_time)


# ####################################### prepare ###########################################
def check_result_path():
    ui.data.real_live_path = ui.data.result_path + ui.data.slash + ui.data.today + ui.data.slash + ui.data. \
        result_live_path
    ui.data.real_live_name = ui.data.real_live_path + ui.data.slash + ui.data.today + '_실시간_'
    ui.data.real_view_like_path = ui.data.result_path + ui.data.slash + ui.data.today + ui.data.slash + ui.data. \
        result_view_like_path
    ui.data.real_view_like_name = ui.data.real_view_like_path + ui.data.slash + ui.data.today + '_조회수+좋아요_'

    if not os.path.exists(ui.data.real_live_path):
        os.makedirs(ui.data.real_live_path)
    if not os.path.exists(ui.data.real_view_like_path):
        os.makedirs(ui.data.real_view_like_path)

    excels_live = []
    for (root, dirs, files) in os.walk(ui.data.real_live_path):
        for file in files:
            if '.xlsx' in file and '_실시간_' in file:
                file = str(file)[str(file).find('간') + 2:str(file).find('.')]
                excels_live.append(int(file))
    if len(excels_live) == 0:
        ui.data.live_index = 0
    else:
        excels_live.sort()
        ui.data.live_index = excels_live[len(excels_live) - 1]

    excels_view_like = []
    for (root, dirs, files) in os.walk(ui.data.real_view_like_path):
        for file in files:
            if '.xlsx' in file and '_조회수+좋아요_' in file:
                file = str(file)[str(file).find('요') + 2:str(file).find('.')]
                excels_view_like.append(int(file))
    if len(excels_view_like) == 0:
        ui.data.view_like_index = 0
    else:
        excels_view_like.sort()
        ui.data.view_like_index = excels_view_like[len(excels_view_like) - 1]


def start_program():
    try:
        ui.data.today = str(datetime.now().date())
        if ui.radio_var.get() == 2:
            # ui.data.first_start = True
            # 입력과 파일 형식 확인
            file_path = valid_file_combo()
            if file_path == '':
                return
            ui.data.channel_list_names, ui.data.channel_list_urls = \
                valid_channel_list_excel_form(file_path)
            if not ui.data.channel_list_names:
                return
            loop_time = int(ui.combo_select_loop_time.get()[:ui.combo_select_loop_time.get().find('분')])
            ui.data.time_term = loop_time
            ui.data.time_sum = 0.0
            ui.data.time_cnt = 0
            # 파일 인덱스 구하기
            check_result_path()
            # 버튼 설정 & 프로그램 실행
            ui.data.stop = False
            ui.button_start.config(state='disabled')
            ui.button_find_file.config(state='disabled')
            ui.combo_select_loop_time.config(state='disabled')

            run_thread = threading.Thread(target=run())
            run_thread.daemon = True

            # 근데 왜 처음 모니터링 부분에서 로딩이 길지 => 1초 정도 차이를 두고 2개의 스레드 실행으로 해결
            # 다른 문제점: 스타트 클릭했을 때만 동작을 하고 스탑 버튼을 클릭했을 때 멈추게 할지
            monitoring_thread = threading.Timer(1, monitoring)
            monitoring_thread.daemon = True

            monitoring_thread.start()
            run_thread.start()

            # thread.join()
            # 스레드의 종료를 기다렸다가 처리되어야 할 때 사용
            # 스레드 안에서 무한루프가 실행되고 있는 상황에서는 조인 사용 x
        else:
            # 월요일
            ui.button_start.config(state='disabled')
            ui.button_stop.config(state='disabled')

            # 월요일 조회수 폴더가 있는지 확인
            ui.data.real_monday_path = ui.data.result_path + ui.data.slash + ui.data.result_monday_path
            ui.data.real_monday_name = ui.data.real_monday_path + ui.data.slash + ui.data.today + '_월요일'
            if not os.path.exists(ui.data.real_monday_path):
                os.makedirs(ui.data.real_monday_path)

            check_result_path()

            monitoring_thread = threading.Thread(target=monday_view())
            monitoring_thread.daemon = True

            monitoring_thread.start()
            monitoring_thread.join()

            ui.button_start.config(state='normal')
            ui.button_stop.config(state='normal')
            ui.progress_var.set(0)
            ui.progress_bar.update()
    except:
        msgbox.showerror("에러", "에러 문의 주세요.")
        err = traceback.format_exc()
        ErrorLog(str(err))
        stop_program()


def stop_program():
    try:
        ui.data.stop = True
        ui.data.run_thread.cancel()
        ui.data.monitoring_thread.cancel()
        ui.progress_var.set(0)
        ui.progress_bar.update()
        ui.button_start.config(state='normal')
        ui.button_find_file.config(state='normal')
        ui.combo_select_loop_time.config(state='readonly')
        print_in_list_box("일시 정지")
    except:
        msgbox.showerror("에러", "에러 문의 주세요.")
        err = traceback.format_exc()
        ErrorLog(str(err))


def test1():
    pass


@dataclass
class DATA:
    stop = False
    today = str(datetime.now().date())
    time_term = 0
    channel_list_names = []
    channel_list_urls = []

    # path, name
    slash = '/'
    result_path = '결과'
    result_live_path = '실시간 시청자 수'
    result_view_like_path = '조회수+좋아요'
    result_monday_path = '월요일 조회수'
    real_live_path = ''
    real_view_like_path = ''
    real_monday_path = ''
    real_live_name = ''
    real_view_like_name = ''
    real_monday_name = ''

    default_live = '필수_실시간.xlsx'
    default_view_like = '필수_조회수+좋아요.xlsx'

    # for live streaming videos
    live_index = 0
    df_live = pd.DataFrame()

    # for ended videos
    view_like_index = 0
    first_start = True
    check_video_list = {}
    df_view_like = pd.DataFrame()

    # for time average
    time_sum = 0.0
    time_cnt = 0

    run_thread = threading.Timer(0, test1)  #
    monitoring_thread = threading.Timer(0, test1)  #


class UI:
    window = tk.Tk()
    window.title('집계 프로그램')
    win_width = 640
    win_height = 400
    window.geometry(str(win_width) + 'x' + str(win_height) + '+100+100')
    window.resizable(False, False)

    excel_ext = r"*.xlsx *.xls *.csv"

    def __init__(self, data):
        self.frame_up, self.frame_down = self.make_frames()
        self.label_search, self.label_combo, self.label_button = self.make_frame_up_labels()
        self.box_find_file, self.button_find_file = self.fill_label_search()
        self.radio_var, self.radio_monday, self.radio_live, self.combo_select_loop_time = self.fill_label_radio_combo()
        self.box_result_comment = self.fill_frame_down()
        self.progress_bar, self.progress_var, self.button_start, self.button_stop = self.fill_label_button()
        self.data = data

    def make_frames(self):
        frame_up = tk.Frame(self.window)  # , relief='solid', bd=1
        frame_up.pack(side='top', fill='both', expand=True)
        frame_down = tk.Frame(self.window)  # , relief='solid', bd=1
        frame_down.pack(side='bottom', fill='both', expand=True)
        return frame_up, frame_down

    def make_frame_up_labels(self):
        label_combo = tk.Label(self.frame_up)  # , relief='solid'
        label_combo.pack(side='top', padx=5, pady=1)  # , fill='both'
        label_search = tk.Label(self.frame_up)
        label_search.pack(side='top', padx=5)
        label_button = tk.Label(self.frame_up)  # , relief='solid'
        label_button.pack(side='top', padx=60)  # , fill='both'
        return label_search, label_combo, label_button

    def fill_label_search(self):
        box_find_file = tk.Entry(self.label_search, width=50)
        box_find_file.pack(side='left', padx=1)
        button_find_file = tk.Button(self.label_search, text='찾기', command=self.get_file,
                                     relief="raised", overrelief="sunken")
        button_find_file.pack(side='right')
        return box_find_file, button_find_file

    def fill_label_radio_combo(self):
        radio_var = tk.IntVar()
        radio1 = tk.Radiobutton(self.label_combo, text='월요일 조회수', value=1, variable=radio_var, command=self.clicked_monday_radio)
        radio2 = tk.Radiobutton(self.label_combo, text='실시간 시청자 수', value=2, variable=radio_var, command=self.clicked_live_radio)
        radio1.grid(column=0, row=0, padx=1)
        radio2.grid(column=1, row=0, padx=1)
        # 반복 시간 콤보 박스
        select_time = [str(i) + "분" for i in range(1, 61)]
        loop_time = ttk.Combobox(self.label_combo, height=10, values=select_time, state='readonly')
        loop_time.set('반복 시간 선택')
        loop_time.grid(column=2, row=0, padx=1)
        radio2.select()
        radio1.config(state='disabled')
        return radio_var, radio1, radio2, loop_time

    def fill_label_button(self):
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(self.label_button, maximum=100, length=157, variable=progress_var)
        progress_bar.grid(column=0, row=0, padx=1)
        # start
        button_start = tk.Button(self.label_button, text='시작', command=start_program,
                                 relief="raised", overrelief="sunken", width=15)
        button_start.grid(column=1, row=0, padx=1)
        # stop
        button_stop = tk.Button(self.label_button, text='정지', command=stop_program,
                                relief="raised", overrelief="sunken", width=15)
        button_stop.grid(column=2, row=0, padx=1)
        return progress_bar, progress_var, button_start, button_stop

    def fill_frame_down(self):
        # 리스트 박스
        box_result_comment = tk.Listbox(self.frame_down, width=88, height=21)  # , font=("Helvetica", 10)
        box_result_comment.pack(side='left', fill='y')
        # 스크롤 바
        scrollbar = tk.Scrollbar(self.frame_down, orient='vertical', width=25)
        scrollbar.config(command=box_result_comment.yview)
        scrollbar.pack(side='right', fill='y')
        # 리스트 박스와 스크롤 바 연결
        box_result_comment.config(yscrollcommand=scrollbar.set)
        return box_result_comment

    def get_file(self):
        file = filedialog.askopenfilenames(filetypes=(("Excel file", self.excel_ext),
                                                      ("all file", "*.*")), initialdir=r"C:\Users")
        self.box_find_file.delete(0, tk.END)
        try:
            self.box_find_file.insert(tk.END, file[0])
        except:
            return

    def clicked_monday_radio(self):
        self.combo_select_loop_time.config(state='disabled')
        self.box_find_file.config(state='disabled')
        self.button_find_file.config(state='disabled')

    def clicked_live_radio(self):
        self.combo_select_loop_time.config(state='readonly')
        self.box_find_file.config(state='normal')
        self.button_find_file.config(state='normal')


def key_input(value):
    print(str(value.keysym))
    if value.keysym == 'Escape':
        exit(0)


def help():
    he = tk.Toplevel(ui.window)
    he.geometry('640x760')
    he.resizable(False, False)
    he.title("설명")
    text = '\
[프로그램 시작 전 필수]\n\n\
1. 집계 프로그램.exe, 필수_실시간.xlsx, 필수_조회수+좋아요.xlsx 세 가지 파일들은 반드시 같은 폴더에 있어야 합니다.\n\n\
2. 채널 엑셀 파일에 필요한 요소\n\
- 첫 번째 행에서 첫 번째 열에는 "채널 이름"이, 두 번째 열에는 "실시간 주소"가 고정되어 있어야 합니다.\n\
- 추가하려는 "채널 이름"과 "실시간 주소"를 짝을 맞춰서 저장해 주세요.\n\
ex) 채널 이름=뭉치의 개팔상팔, 실시간 주소=https://www.youtube.com/@whyrano_gaemungchi/streams\n\
- 채널_목록.xlsx 파일을 참고하면 됩니다.\n\n\n\
["실시간 시청자 수" 선택 시]\n\n\
1. 실시간 시청자 수 데이터를 실시간으로 저장하고, 실시간 스트리밍이 끝난 영상에 대한 조회수+좋아요 데이터도 실시간으로 저장됩니다.\n\n\
2. 실행 과정 및 결과\n\
- 반복 시간을 선택 후, 채널 정보가 들어간 엑셀 파일을 선택해 주세요.\n\
- 선택을 하고 시작 버튼을 누르면 버튼 왼쪽에 실시간 시청자 수 데이터 수집이 현재 얼마큼 실행되고 있는지 볼 수 있습니다.\n\
- 한 턴이 끝나면 하단에 어떤 파일이 생성되었는지 확인 메시지가 뜹니다.\n\
- 실시간 스트리밍이 종료된 영상이 생기는 경우, 조회수+좋아요 데이터를 수집 및 저장 후 확인 메시지가 뜹니다.\n\n\
3. 결과가 저장된 엑셀\n\
- 실시간 시청자 수\n\
\t - 선택한 반복 시간 단위로 시청자 수가 저장됩니다.\n\
- 조회수+좋아요\n\
\t - 종료된 영상의 조회 수와 좋아요 수가 저장됩니다.\n\
\t - 종료된 영상의 조회 수 또는 좋아요 수 정보가 없으면 엑셀에 "정보 없음"이라 저장됩니다.\n\
\t - 종료된 영상이 삭제된 영상이면 엑셀에 "삭제된 영상"이랑 저장됩니다.\n\n\
4. 결과 저장 위치\n\
- 프로그램과 같은 경로에 "결과" 폴더가 생성되고, 그 안에 "현재 날짜" 폴더가 생성됩니다.\n\
- 실시간 시청자 수 데이터는 "실시간 시청자 수" 폴더에 조회수+좋아요 데이터는 "조회수+좋아요" 폴더에 각각 저장됩니다.\n\
- 한 턴 당 엑셀 파일이 추가로 생성되며, 데이터 누적 저장 시 파일 손상을 줄이기 위함이므로 가장 마지막에 생성된 엑셀만 봐도 됩니다.\n\n\
5. 정지 버튼 클릭 시\n\
- 실시간 시청자 수 데이터와 조회수+좋아요 데이터 수집 및 저장이 중지됩니다.\n\
- 실시간 시청자 수 데이터 수집뿐만 아니라 조회수+좋아요 데이터 수집이 중지되므로 종료된 영상의 조회수+좋아요 데이터가 모두 저장되면 정지 버튼을 누르는 것을 권장합니다.\n\n\n\
[추가]\n\n\
- 필수 엑셀 파일이 없거나 엑셀 형식이 틀리거나 채널의 데이터를 수집할 수 없을 경우 프로그램 하단 부분에 생성되는 메시지를 확인해 주세요.\n\
- "에러 문의해 주세요." 메시지 창이 뜨면 프로그램과 같은 폴더 경로에 생긴 result 파일을 첨부하여 바로 문의해 주세요.'
    # lb = tk.Text(he)
    # lb.insert(tk.INSERT, text)
    lb = tk.Label(he, text=text, wraplength=640, justify='left')
    lb.pack()


def ErrorLog(error: str):
    current_time = time.strftime("%Y.%m.%d/%H:%M:%S", time.localtime(time.time()))
    with open("result.log", "a") as f:
        f.write(f"[{current_time}] - {error}\n")


if __name__ == '__main__':
    try:
        ui = UI(DATA())

        menubar = tk.Menu(ui.window)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="설명", command=help)
        menubar.add_cascade(label="도움말", menu=helpmenu)

        ui.window.config(menu=menubar)

        ui.window.bind('<Key>', key_input)

        ui.window.mainloop()
    except:
        err = traceback.format_exc()
        ErrorLog(str(err))
        stop_program()
