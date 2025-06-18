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


import sys
sys.coinit_flags = 0 #Temporary fix for: Win32 exception occurred releasing IUnknown
from pythoncom import CoInitialize, CoUninitialize
from gc import collect
from os import environ
from os.path import splitdrive
from wmi import WMI

class Util:
    @staticmethod
    def get_user() -> str:
        CoInitialize()
        try:
            wmi_instance = WMI()
            for computer_system in wmi_instance.query("SELECT UserName FROM Win32_ComputerSystem"):
                if computer_system.UserName:
                    username = computer_system.UserName.split('\\')
                    if len(username) == 2:
                        return username[1]
                    return username[0]
            return ""
        finally:
            collect()
            CoUninitialize()
    
    @staticmethod
    def get_download_folder() -> str:
        root_path = splitdrive(environ['SystemRoot'])[0]
        return f'{root_path}\\Users\\{Util.get_user()}\\Downloads'
