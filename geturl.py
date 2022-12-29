import tkinter as tk
import tkinter.ttk as ttk
from dataclasses import dataclass
from tkinter import filedialog
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
        print_in_list_box('default_live.xlsx 파일 위치를 다시 확인해 주세요.')
        return ''
    if not os.path.exists(ui.data.default_view_like):
        print_in_list_box('default_view_like.xlsx 파일 위치를 다시 확인해 주세요.')
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
        names = list(rd_df['name'])
        urls = list(rd_df['url'])
    except:
        print_in_list_box('엑셀 파일 형식을 맞춰 주세요.')
        return [], [], []
    return names, urls


def valid_channel_url_form(name, url):
    if not url.startswith('https://www.youtube.com/') or not url.endswith('/streams'):
        print_in_list_box(str(name) + ': URL 형식 확인해 주세요.')
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
        if data.loc[i]['채널명'] == ch_name and data.loc[i]['영상 아이디'] == video_id:
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
        print_in_list_box(str(name) + ': 해당 페이지를 열 수 없습니다.')
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
    if ui.data.stop:
        return
    start = time.time()
    ui.data.df_live = open_live_excel(ui.data)
    print("====================== START ======================")
    curr_time = get_curr_time()
    ui.data.df_live[curr_time] = 0
    for x in tqdm(range(len(ui.data.channel_list_urls)), leave=True):
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


# ####################################### program ###########################################

# #################################  monitoring program #####################################
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
    row_index = len(ui.data.df_view_like)
    ui.data.df_view_like.loc[row_index, '채널명'] = video_info[0]
    ui.data.df_view_like.loc[row_index, '영상 제목'] = video_info[1]
    ui.data.df_view_like.loc[row_index, '영상 아이디'] = video_info[2]
    if view_cnt == -1:
        ui.data.df_view_like.loc[row_index, '조회수'] = '정보없음'
    else:
        ui.data.df_view_like.loc[row_index, '조회수'] = view_cnt
    if like_cnt == -1:
        ui.data.df_view_like.loc[row_index, '좋아요 수'] = '정보없음'
    else:
        ui.data.df_view_like.loc[row_index, '좋아요 수'] = like_cnt
    return True


def monitoring():
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
                                      backup_df_live.loc[i]['영상 아이디'])] = False
            continue
        if find_video_row_index(backup_df_live.loc[i]['채널명'], backup_df_live.loc[i]['영상 아이디'],
                                ui.data.df_view_like) == -1:
            ui.data.check_video_list[(backup_df_live.loc[i]['채널명'], backup_df_live.loc[i]['영상 제목'],
                                      backup_df_live.loc[i]['영상 아이디'])] = False
    # 크롤링하여 리스트에 저장된 채널이 종료되었는지 확인
    flag = False
    for key, value in ui.data.check_video_list.items():
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

    monitoring_thread = threading.Timer(0, monitoring)
    monitoring_thread.daemon = True
    monitoring_thread.start()


# #################################  monitoring program #####################################

# ####################################### prepare ###########################################
def check_result_path():
    ui.data.real_live_path = ui.data.result_path + ui.data.slash + ui.data.today + ui.data.slash + ui.data. \
        result_live_path
    ui.data.real_live_name = ui.data.real_live_path + ui.data.slash + ui.data.today + '_실시간_'
    if not os.path.exists(ui.data.real_live_path):
        os.makedirs(ui.data.real_live_path)
    if os.path.exists(ui.data.real_live_name + '1.xlsx'):
        ui.data.live_index = 2
        while True:
            if not os.path.exists(ui.data.real_live_name + str(ui.data.live_index) + '.xlsx'):
                ui.data.live_index = ui.data.live_index - 1
                break
            else:
                ui.data.live_index = ui.data.live_index + 1

    # view_like_index
    ui.data.real_view_like_path = ui.data.result_path + ui.data.slash + ui.data.today + ui.data.slash + ui.data. \
        result_view_like_path
    ui.data.real_view_like_name = ui.data.real_view_like_path + ui.data.slash + ui.data.today + '_조회수+좋아요_'
    if not os.path.exists(ui.data.real_view_like_path):
        os.makedirs(ui.data.real_view_like_path)
    if os.path.exists(ui.data.real_view_like_name + '1.xlsx'):
        ui.data.view_like_index = 2
        while True:
            if not os.path.exists(ui.data.real_view_like_name + str(ui.data.view_like_index) + '.xlsx'):
                ui.data.view_like_index = ui.data.view_like_index - 1
                break
            else:
                ui.data.view_like_index = ui.data.view_like_index + 1


def start_program():
    ui.data.first_start = True
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
    ui.data.today = str(datetime.now().date())
    check_result_path()
    # 버튼 설정 & 프로그램 실행
    ui.data.stop = False
    ui.button_start.config(state='disabled')
    ui.button_find_file.config(state='disabled')
    ui.combo_select_loop_time.config(state='disabled')

    run_thread = threading.Thread(target=run())
    run_thread.daemon = True

    # 근데 왜 처음 모니터링 부분에서 로딩이 길지 => 1초 정도 차이를 두고 2개의 스레드 실행으로 해결
    monitoring_thread = threading.Timer(1, monitoring)
    monitoring_thread.daemon = True

    monitoring_thread.start()
    run_thread.start()

    # thread.join()
    # 스레드의 종료를 기다렸다가 처리되어야 할 때 사용
    # 스레드 안에서 무한루프가 실행되고 있는 상황에서는 조인 사용 x


def stop_program():
    ui.data.stop = True
    ui.data.run_thread.cancel()
    ui.progress_var.set(0)
    ui.progress_bar.update()
    ui.button_start.config(state='normal')
    ui.button_find_file.config(state='normal')
    ui.combo_select_loop_time.config(state='normal')
    print_in_list_box("일시 정지")


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
    real_live_path = ''
    real_view_like_path = ''
    real_live_name = ''
    real_view_like_name = ''

    default_live = 'default_live.xlsx'
    default_view_like = 'default_view_like.xlsx'

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

    run_thread = threading.Thread()


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
        self.combo_select_goal, self.combo_select_loop_time = self.fill_label_combo()
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
        label_search = tk.Label(self.frame_up)
        label_search.pack(side='top', padx=5)
        label_combo = tk.Label(self.frame_up)  # , relief='solid'
        label_combo.pack(side='top', padx=5, pady=1)  # , fill='both'
        label_button = tk.Label(self.frame_up)  # , relief='solid'
        label_button.pack(side='top', padx=60)  # , fill='both'
        return label_search, label_combo, label_button

    def fill_label_search(self):
        box_find_file = tk.Entry(self.label_search, width=41)
        box_find_file.pack(side='left', padx=1)
        button_find_file = tk.Button(self.label_search, text='찾기', command=self.get_file,
                                     relief="raised", overrelief="sunken")
        button_find_file.pack(side='right')
        return box_find_file, button_find_file

    def fill_label_combo(self):
        select_goal = ['실시간 시청자 수', '조회수']
        combo1 = ttk.Combobox(self.label_combo, height=5, values=select_goal, state='readonly')
        combo1.set('실시간 시청자 수')
        combo1.grid(column=0, row=0, padx=1)
        combo1.config(state='disabled')
        # 반복 시간 콤보 박스
        select_time = [str(i) + "분" for i in range(1, 61)]
        loop_time = ttk.Combobox(self.label_combo, height=10, values=select_time, state='readonly')
        loop_time.set('반복 시간 선택')
        loop_time.grid(column=1, row=0, padx=1)
        return combo1, loop_time

    def fill_label_button(self):
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(self.label_button, maximum=100, length=150, variable=progress_var)
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


def key_input(value):
    if value.keysym == 'Escape':
        exit(0)


if __name__ == '__main__':
    ui = UI(DATA())

    ui.window.bind('<Key>', key_input)

    ui.window.mainloop()
