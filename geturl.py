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


def open_live_view_like_excel(data):
    if data.live_index != 0:
        df_live = pd.read_excel(ui.data.real_live_name + str(data.live_index) + '.xlsx')
    else:
        df_live = pd.read_excel(ui.data.default_live)

    if data.view_like_index != 0:
        df_view_like = pd.read_excel(ui.data.real_view_like_name + str(data.view_like_index) + '.xlsx')
    else:
        df_view_like = pd.read_excel(ui.data.default_view_like)
    return df_live, df_view_like


def find_video_row_index(ch_name, video_id, data):
    for i in range(len(data)):
        if data.loc[i]['채널명'] == ch_name and data.loc[i]['영상 아이디'] == video_id:
            return i
    return -1


def fill_check_video_list():
    if ui.data.first_start:
        ui.data.first_start = False
        ui.data.check_video_list = {}
        for i in range(len(ui.data.df_live)):
            # if find_video(df_live.loc[i]['채널명'], df_live.loc[i]['영상 아이디'], df_view_like) == -1:
            ui.data.check_video_list[(ui.data.df_live.loc[i]['채널명'], ui.data.df_live.loc[i]['영상 제목'],
                                      ui.data.df_live.loc[i]['영상 아이디'])] = False
    for key, value in ui.data.check_video_list.items():
        ui.data.check_video_list[key] = False


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
                            video_title = (video.get('title').get('runs'))[0].get('text')
                            viewers = (video.get('viewCountText').get('runs'))[0].get('text')
                            # update video status
                            dic_keys = list(data.check_video_list.keys())
                            for key in dic_keys:
                                if key[2] == href:
                                    if key[1] != video_title:
                                        data.check_video_list.pop(key)
                                    break
                            data.check_video_list[(channel_name, video_title, href)] = True
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
    # return data.df_live


def view_like_scraping(video_info, parsing):
    # f = open('test.txt', 'w', newline='', encoding='utf-16')
    f = open('test.txt', 'a', newline='', encoding='utf-16')
    for line in parsing:
        if 'ytInitialData' in line:
            line = line.strip('var ytInitialData = ').strip(';')
            # line = line.strip(';')
            json_object = json.loads(line)
            contents = json_object.get('contents').get('twoColumnWatchNextResults').get('results').get('results').get(
                'contents')
            for content in contents:
                if '좋아요' in str(content):
                    # 제목
                    f.write(video_info[1] + '\n\n\n\n')
                    f.write(video_info[2] + '\n\n\n\n')
                    content = content.get('videoPrimaryInfoRenderer')
                    # 조회수
                    views = content.get('viewCount').get('videoViewCountRenderer').get('viewCount')
                    if '조회수' not in str(views):
                        break
                    views = views.get('simpleText')
                    f.write(str(views) + '\n\n\n\n')
                    view_cnt = str(views)[4:]
                    view_cnt = int(view_cnt[:view_cnt.find('회')].replace(',', ''))
                    f.write(str(view_cnt) + '\n\n\n\n')
                    items = content.get('videoActions').get('menuRenderer').get('topLevelButtons')
                    for like_cnt in items:
                        if '좋아요' in str(like_cnt):
                            # 좋아요 수
                            like_cnt = like_cnt.get('segmentedLikeDislikeButtonRenderer').get('likeButton').get(
                                'toggleButtonRenderer').get('defaultText').get('accessibility').get(
                                'accessibilityData').get('label')
                            like_cnt = int(str(like_cnt)[4:str(like_cnt).find('개')].replace(',', ''))
                            f.write(str(like_cnt) + '\n\n\n\n')
                            row_index = len(ui.data.df_view_like)
                            ui.data.df_view_like.loc[row_index, '채널명'] = video_info[0]
                            ui.data.df_view_like.loc[row_index, '영상 제목'] = video_info[1]
                            ui.data.df_view_like.loc[row_index, '영상 아이디'] = video_info[2]
                            ui.data.df_view_like.loc[row_index, '조회수'] = view_cnt
                            ui.data.df_view_like.loc[row_index, '좋아요 수'] = like_cnt
                            # view_like도 인덱스 만들기
                        break
                    break
            f.close()
            break


def manage_ended_videos():
    f = open('test.txt', 'w', newline='', encoding='utf-16')
    f.close()
    end_video = []
    for key, value in ui.data.check_video_list.items():
        if not value:
            end_video.append(key)
    if len(end_video) == 0:
        return
    ui.data.view_like_index = ui.data.view_like_index + 1
    for e in end_video:
        print(e)
        parsing = crawling(e[0] + '(' + e[1] + ')', 'https://youtube.com/' + e[2])
        if parsing is None:  # 페이지가 지워졌을 때는 엑셀에 어떻게 표시할지 생각해보자 - 그냥 비워 둘까
            continue
        view_like_scraping(e, parsing)
        ui.data.df_view_like.to_excel(ui.data.real_view_like_name + str(ui.data.view_like_index) + '.xlsx', index=False)
        ui.data.check_video_list.pop(e)


def run():
    if ui.data.stop:
        return
    start = time.time()
    ui.data.df_live, ui.data.df_view_like = open_live_view_like_excel(ui.data)
    fill_check_video_list()
    print("====================== START ======================")
    curr_time = get_curr_time()
    ui.data.df_live[curr_time] = 0
    for x in tqdm(range(len(ui.data.channel_list_urls))):
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
    print_in_list_box(ui.data.today + '_' + str(ui.data.live_index) + '.xlsx')
    print('총 실행 시간: %s 초' % (time.time() - start))
    print("======================= END =======================")
    # manage_ended_videos()
    threading.Timer(0, manage_ended_videos).start()

    global run_thread
    timer = (60 * int(ui.data.time_term) - (len(ui.data.channel_list_names) // 2) + 2)
    run_thread = threading.Timer(timer, run)
    run_thread.start()


# ####################################### program ###########################################


# ####################################### prepare ###########################################
def check_result_path():
    ui.data.real_live_path = ui.data.result_path + ui.data.slash + ui.data.today + ui.data.slash + ui.data.\
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
    ui.data.real_view_like_path = ui.data.result_path + ui.data.slash + ui.data.today + ui.data.slash + ui.data.\
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
    # 파일 인덱스 구하기
    ui.data.today = str(datetime.now().date())
    check_result_path()
    # 버튼 설정 & 프로그램 실행
    ui.data.stop = False
    ui.button_start.config(state='disabled')
    ui.button_find_file.config(state='disabled')
    ui.combo_select_loop_time.config(state='disabled')

    threading.Timer(0, run).start()


def stop_program():
    ui.data.stop = True
    run_thread.cancel()
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
        # start_btn.pack(side='left')
        # stop_btn.pack(side='right')

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


def test():
    return


run_thread = threading.Timer(600, test)

if __name__ == '__main__':
    ui = UI(DATA())

    ui.window.bind('<Key>', key_input)

    ui.window.mainloop()
