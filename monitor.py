from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
import logging
import time
from random import randrange as rr
from all_update import update
import os
import glob
from utilities import check_expired, remove_temp

SLASH = '/'
NOT_SLASH = '\\'
EXCEL_FILE = 'Axa Marval Status Report.xlsx'
EXCEL_LIST = []

def call_update(file_path, directory):
    print(file_path+"======"+directory)
    update(file_path, directory)
    #update_axa_marvel(file_)
    #update_axa_sp_infocity(BASE_DIR, SLASH, EXCEL_FILE)

class Event(LoggingEventHandler):
    def __init__(self, file_path, ui_handler):
      self.file_path = file_path
      self.ui_handler = ui_handler

    def dispatch(self, event):
        global EXCEL_LIST
        print event.event_type, event.src_path, EXCEL_LIST
        if check_expired():
            self.ui_handler.expired_view()
        elif event.event_type in ['created', 'modified'] and event.src_path in (EXCEL_LIST):
            x = str(rr(1000))
            print 'gottcha - '+x
            time.sleep(1)
            try:
                call_update(self.file_path, BASE_DIR)
            except Exception as e:
                print e
            print 'done - '+x

def main(directory, ui_handler):
    global BASE_DIR, FILE_PATH, EXCEL_LIST
    if directory[-1] != SLASH:
        directory += SLASH
    BASE_DIR = directory
    print "2******************"
    print BASE_DIR
    xcel_list = glob.glob(BASE_DIR+"/*.xlsx")
    print xcel_list
    xcel_list = remove_temp(xcel_list)
    EXCEL_LIST = xcel_list
	
    for i in xcel_list:

        call_update(i, directory)
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
        event_handler = Event(i, ui_handler)
        observer = Observer()
        observer.schedule(event_handler, BASE_DIR, recursive=True)
        observer.start()