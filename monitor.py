from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
import logging
import time
from random import randrange as rr
from axa_marvel import update_axa_marvel
from axa_sp_infocity import update_axa_sp_infocity
from all_update import update
import os
import glob

SLASH = '/'
NOT_SLASH = '\\'
EXCEL_FILE = 'Axa Marval Status Report.xlsx'

def call_update(file_path, directory):
    print(file_path+"======"+directory)
    update(file_path, directory)
    #update_axa_marvel(file_)
    #update_axa_sp_infocity(BASE_DIR, SLASH, EXCEL_FILE)

class Event(LoggingEventHandler):
    def __init__(self, file_path):
      self.file_path = file_path
      
    def dispatch(self, event):
        # print 'event occured', event.event_type, event.src_path
        if event.event_type in ['created', 'modified'] and event.src_path == (self.file_path):
            x = str(rr(1000))
            print 'gottcha - '+x
            time.sleep(1)
            call_update(self.file_path, BASE_DIR)
            print 'done - '+x

def main(directory):
    global BASE_DIR, FILE_PATH
    if directory[-1] != SLASH:
        directory += SLASH
    BASE_DIR = directory
    xcel_list = glob.glob(BASE_DIR+"/*.xlsx")

    for i in xcel_list:
        print i

        call_update(i, directory)
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
        event_handler = Event(i)
        observer = Observer()
        observer.schedule(event_handler, BASE_DIR, recursive=True)
        observer.start()



    '''call_update(BASE_DIR, SLASH, EXCEL_FILE)
    logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
    event_handler = Event()
    observer = Observer()
    observer.schedule(event_handler, BASE_DIR, recursive=True)
    observer.start()'''
