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

from os.path import exists, isfile, normpath
from platform import architecture
from string import ascii_uppercase

class Extractor:

    @staticmethod
    def get_extractor() -> ( dict | None):
        arch: str = architecture()[0]
        drives: list = []
        extractors: dict = {}

        for drive in ascii_uppercase:
            if exists(f"{drive}:\\"):
                drives.append(drive)

        match arch:
            case '32bit':
                paths = {
                    'UnRAR': f'{{drive}}:\\Program Files (x86)\\WinRAR\\UnRAR.exe',
                    '7Zip': f'{{drive}}:\\Program Files (x86)\\7-Zip\\7z.exe'
                }

            case '64bit':
                paths = {
                    'UnRAR': f'{{drive}}:\\Program Files\\WinRAR\\UnRAR.exe',
                    '7Zip': f'{{drive}}:\\Program Files\\7-Zip\\7z.exe'
                }

            case _: 
                return None
        
        for drive in drives:
            for extractor, path in paths.items():
                if isfile(path.format(drive=drive)):
                    extractors[extractor] = path.format(drive=drive)
        return extractors if extractors else None


class Zip:

    @staticmethod
    def is_zip_file(file_path: str) -> bool:
        return Zip.is_normal_zip_file(file_path) or Zip.is_spanned_zip_file(file_path) or file_path.endswith('.zip')

    @classmethod
    def is_normal_zip_file(cls, file_path: str) -> bool:
        try:
            with open(file_path, 'rb') as file:
                return file.read(4) == b'PK\x03\x04'
        except (FileNotFoundError, IOError):
            return False
    
    
    @classmethod
    def is_spanned_zip_file(cls, file_path: str) -> bool:
        try:
            with open(file_path, 'rb') as file:
                return file.read(4) == b'PK\x07\x08'
        except (FileNotFoundError, IOError):
            return False

class Rar:
    
    @staticmethod
    def is_rar_file( file_path: str) -> bool:
        return Rar.is_old_rar_file(file_path) or Rar.is_new_rar_file(file_path) or file_path.endswith('.rar')

    @classmethod
    def is_old_rar_file(cls, file_path: str) -> bool:
        try:
            with open(normpath(file_path), 'rb') as file:
                return file.read(7) == b'Rar!\x1A\x07\x00'
        except (FileNotFoundError, IOError):
            return False

    @classmethod
    def is_new_rar_file(cls, file_path: str) -> bool:
        try:
            with open(normpath(file_path), 'rb') as file:
                return file.read(8) == b'Rar!\x1A\x07\x01\x00'
        except (FileNotFoundError, IOError):
            return False
