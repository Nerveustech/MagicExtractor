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


from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from Files import Zip, Rar, Extractor
from gc import collect, enable, disable
from os import makedirs, remove
from os.path import dirname, basename, exists, join, isfile
from servicemanager import Initialize, PrepareToHostSingle, StartServiceCtrlDispatcher, LogMsg, EVENTLOG_INFORMATION_TYPE, PYS_SERVICE_STARTED, PYS_SERVICE_STOPPED
from socket import setdefaulttimeout
from subprocess import Popen
from sys import argv, executable
from threading import Timer
from Util import Util
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
from win32event import CreateEvent, SetEvent, WaitForSingleObject, WAIT_OBJECT_0
from win32service import SERVICE_STOP_PENDING, SERVICE_RUNNING
from win32serviceutil import ServiceFramework, HandleCommandLine
from zipfile import ZipFile


@dataclass
class MagicExtractor_Config:
    program_path: str
    user_download_folder: str
    extractors: dict


config = MagicExtractor_Config(program_path=dirname(executable), user_download_folder=Util.get_download_folder(),
                               extractors=Extractor.get_extractor())

class MagicExtractor:

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
            folder_to_extract = MagicExtractor.get_incremented_directory_name(folder_to_extract)
            makedirs(folder_to_extract)
        
        with ZipFile(file_path) as unzipit:
            unzipit.extractall(folder_to_extract)
        MagicExtractor.send_file_to_recycle_bin(file_path)
    
    @staticmethod
    def unrar_file(file_path: str, folder_to_extract: str) -> None:
        if not exists(folder_to_extract):
            makedirs(folder_to_extract)
        else:
            folder_to_extract = MagicExtractor.get_incremented_directory_name(folder_to_extract)
            makedirs(folder_to_extract)
        
        process = Popen([config.extractors['UnRAR'], 'x' , file_path, folder_to_extract])
        process.wait()

        MagicExtractor.send_file_to_recycle_bin(file_path)
    
    @staticmethod
    def decompress_files_concurrently(files: list) -> None:
        with ThreadPoolExecutor(max_workers=4) as executor:
            for file_path in files:
                if file_path.endswith('.zip'):
                    folder_to_extract = file_path.removesuffix('.zip')
                    executor.submit(MagicExtractor.unzip_file, file_path, folder_to_extract)
                elif file_path.endswith('.rar'):
                    folder_to_extract = file_path.removesuffix('.rar')
                    executor.submit(MagicExtractor.unrar_file, file_path, folder_to_extract)

    
    def send_file_to_recycle_bin(file_path: str) -> None:
        if isfile(file_path) and file_path.endswith('.zip') or file_path.endswith('.rar'):
            remove(file_path)
       

class ZipFileWatcher(FileSystemEventHandler):
    def __init__(self):
            super().__init__()
            self.files_to_process = {
                'zip': {'files': [], 'timer': None, 'decompress_method': self.process_zip_files},
                'rar': {'files': [], 'timer': None, 'decompress_method': self.process_rar_files}
            }

    def on_created(self, event) -> None:
        if event.is_directory:
            return None

        file_extension = event.src_path.split('.')[-1].lower()
        if file_extension in self.files_to_process:
            if self.is_valid_file(file_extension, event.src_path):
                print(file_extension, event.src_path)
                self.add_file(file_extension, event.src_path)
                self.reset_timer(file_extension)


    def is_valid_file(self, file_extension: str, file_path: str) -> bool:
        match file_extension:
            case 'zip':
                return Zip.is_zip_file(file_path)
            case 'rar':
                return Rar.is_rar_file(file_path)
            case _:
                return False

    def add_file(self, file_extension, file_path) -> None:
        self.files_to_process[file_extension]['files'].append(file_path)


    def reset_timer(self, file_extension) -> None:
        timer = self.files_to_process[file_extension]['timer']
        if timer is not None:
            timer.cancel()
        self.files_to_process[file_extension]['timer'] = Timer(1.0, self.files_to_process[file_extension]['decompress_method'])
        self.files_to_process[file_extension]['timer'].start()

    def process_zip_files(self) -> None:
        self.process_files('zip')

    def process_rar_files(self) -> None:
        self.process_files('rar')

    def process_files(self, file_extension) -> None:
        files = self.files_to_process[file_extension]['files']
        if files:
            MagicExtractor.decompress_files_concurrently(files)
            files.clear()
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
        self.folder_downloads = config.user_download_folder
        setdefaulttimeout(60)
        enable()

    def SvcStop(self):
        self.ReportServiceStatus(SERVICE_STOP_PENDING)
        LogMsg(EVENTLOG_INFORMATION_TYPE, PYS_SERVICE_STOPPED, (self._svc_name_, ''))
        SetEvent(self.stop_event)
        self.observer.stop()
        self.observer.join()
        disable()
    
    def SvcDoRun(self):
        self.ReportServiceStatus(SERVICE_RUNNING)
        LogMsg(EVENTLOG_INFORMATION_TYPE, PYS_SERVICE_STARTED, (self._svc_name_, ''))
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