#
# MIT License
#
# Copyright (c) 2025 Andrea Michael M. Molino
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#

from os import environ, system, makedirs, remove
from os.path import splitdrive, dirname, basename, exists, join, isfile
from pythoncom import CoInitialize, CoUninitialize
from servicemanager import Initialize, PrepareToHostSingle, StartServiceCtrlDispatcher, LogMsg, EVENTLOG_INFORMATION_TYPE, PYS_SERVICE_STARTED
from sys import argv, executable
from time import sleep
from subprocess import Popen
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
from win32event import CreateEvent, SetEvent, WaitForSingleObject, WAIT_OBJECT_0
from win32service import SERVICE_STOP_PENDING, SERVICE_RUNNING
from win32serviceutil import ServiceFramework, HandleCommandLine
from wmi import WMI
from zipfile import ZipFile
from gc import collect, enable, disable, DEBUG_LEAK

from concurrent.futures import ThreadPoolExecutor
from threading import Timer

class Util:
    
    @staticmethod
    def get_download_folder() -> str:
        CoInitialize()
        try:
            wmi_instance = WMI()

            for computer_system in wmi_instance.query("SELECT UserName FROM Win32_ComputerSystem"):
                username = str(computer_system.UserName).split('\\')[1]
        
            root_path = splitdrive(environ['SystemRoot'])[0]
            return f'{root_path}\\Users\\{username}\\Downloads'
        finally:
            CoUninitialize()

    @staticmethod
    def get_incremented_directory_name(folder: str) -> str:
        parent_directory = dirname(folder)
        directory_name = basename(folder)

        index = 2
        directory_to_increment = folder

        while exists(directory_to_increment):
            directory_to_increment = join(parent_directory, f"{directory_name} ({index})")
            index += 1

        return directory_to_increment
    
    @staticmethod
    def unzip_file(file_path: str, folder_to_extract: str) -> None:
        if not exists(folder_to_extract):
            makedirs(folder_to_extract)
        else:
            folder_to_extract = Util.get_incremented_directory_name(folder_to_extract)
            makedirs(folder_to_extract)
        
        with ZipFile(file_path) as unzipit:
            unzipit.extractall(folder_to_extract)
        Util.send_file_to_recycle_bin(file_path)
    
    @staticmethod
    def unzip_files_concurrently(files: list) -> None:
        with ThreadPoolExecutor(max_workers=4) as executor:
            for file_path in files:
                folder_to_extract = file_path.removesuffix('.zip')
                executor.submit(Util.unzip_file, file_path, folder_to_extract)
    
    def send_file_to_recycle_bin(file_path: str):
        if isfile(file_path) and file_path.endswith('.zip'):
            remove(file_path)
    

        

class ZipFileWatcher(FileSystemEventHandler):
    def __init__(self):
            super().__init__()
            self.zip_files = []
            self.timer = None

    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith('.zip'):
            self.zip_files.append(event.src_path)
            self.reset_timer()

    def reset_timer(self):
        if self.timer is not None:
            self.timer.cancel()
        self.timer = Timer(1.0, self.process_zip_files)
        self.timer.start()

    def process_zip_files(self):
        if self.zip_files:
            Util.unzip_files_concurrently(self.zip_files)
            self.zip_files.clear()
            collect()


class Worker(ServiceFramework):
    _svc_name_ = "MagicExtractor.Service"
    _svc_display_name_ = "MagicExtractor Service"
    _svc_description_ = ""

    def __init__(self, args):
        ServiceFramework.__init__(self, args)
        self.stop_event = CreateEvent(None, 0, 0, None)
        self.event_handler = ZipFileWatcher()
        self.observer = Observer()
        self.folder_downloads = Util.get_download_folder()
        enable()

    def SvcStop(self):
        self.ReportServiceStatus(SERVICE_STOP_PENDING)
        SetEvent(self.stop_event)
        self.observer.stop()
        self.observer.join()
        disable()
    
    def SvcDoRun(self):
        self.ReportServiceStatus(SERVICE_RUNNING)
        #LogMsg(EVENTLOG_INFORMATION_TYPE, PYS_SERVICE_STARTED, (self._svc_name_, ''))
        self.observer.schedule(self.event_handler, self.folder_downloads, recursive=False)
        self.observer.start()

        try:
            while True:
                result = WaitForSingleObject(self.stop_event, 1000)
                if result == WAIT_OBJECT_0:
                    break

        except Exception as e:
            self.observer.stop()

        finally:
            self.observer.stop()
            self.observer.join()


if __name__ == '__main__':
    if len(argv) == 1:
        Initialize()
        PrepareToHostSingle(Worker)
        StartServiceCtrlDispatcher()
    else:
        HandleCommandLine(Worker)