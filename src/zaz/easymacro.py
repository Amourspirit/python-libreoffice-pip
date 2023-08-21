#!/usr/bin/env python3

# == Rapid Develop Macros in LibreOffice ==

# ~ This file is part of ZAZ.

# ~ https://git.cuates.net/elmau/zaz

# ~ ZAZ is free software: you can redistribute it and/or modify
# ~ it under the terms of the GNU General Public License as published by
# ~ the Free Software Foundation, either version 3 of the License, or
# ~ (at your option) any later version.

# ~ ZAZ is distributed in the hope that it will be useful,
# ~ but WITHOUT ANY WARRANTY; without even the implied warranty of
# ~ MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# ~ GNU General Public License for more details.

# ~ You should have received a copy of the GNU General Public License
# ~ along with ZAZ.  If not, see <https://www.gnu.org/licenses/>.

import base64
import csv
import ctypes
import datetime
import getpass
import gettext
import hashlib
import io
import json
import logging
import os
import platform
import re
import shlex
import shutil
import socket
import ssl
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import zipfile

from collections import OrderedDict
from collections.abc import MutableMapping
from decimal import Decimal
from enum import IntEnum
from functools import wraps
from pathlib import Path
from pprint import pprint
from socket import timeout
from string import Template
from typing import Any, Union
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

import imaplib
import smtplib
from smtplib import SMTPException, SMTPAuthenticationError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import mailbox

import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt.MessageBoxResults import YES
from com.sun.star.awt import Rectangle, Size, Point
from com.sun.star.awt.PosSize import POSSIZE, SIZE
from com.sun.star.awt import Key, KeyModifier, KeyEvent
from com.sun.star.container import NoSuchElementException
from com.sun.star.datatransfer import XTransferable, DataFlavor

from com.sun.star.beans import PropertyValue, NamedValue
from com.sun.star.sheet import TableFilterField
from com.sun.star.table.CellContentType import EMPTY, VALUE, TEXT, FORMULA
from com.sun.star.util import Time, Date, DateTime

from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

from com.sun.star.lang import Locale
from com.sun.star.lang import XEventListener
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XMenuListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import XMouseMotionListener
from com.sun.star.awt import XFocusListener
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import XItemListener
from com.sun.star.awt import XTabListener
from com.sun.star.awt import XWindowListener
from com.sun.star.awt import XTopWindowListener
from com.sun.star.awt.grid import XGridDataListener
from com.sun.star.awt.grid import XGridSelectionListener
from com.sun.star.script import ScriptEventDescriptor

from com.sun.star.io import IOException, XOutputStream

# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1FontUnderline.html
from com.sun.star.awt import FontUnderline
from com.sun.star.style.VerticalAlignment import TOP, MIDDLE, BOTTOM

from com.sun.star.view.SelectionType import SINGLE, MULTI, RANGE

from com.sun.star.sdb.CommandType import TABLE, QUERY, COMMAND

try:
    from peewee import Database, DateTimeField, DateField, TimeField, __exception_wrapper__
except ImportError as e:
    Database = DateField = TimeField = DateTimeField = object
    print("You need install peewee, only if you will develop with Base")


LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
LOG_DATE = "%d/%m/%Y %H:%M:%S"
logging.addLevelName(logging.ERROR, "\033[1;41mERROR\033[1;0m")
logging.addLevelName(logging.DEBUG, "\x1b[33mDEBUG\033[1;0m")
logging.addLevelName(logging.INFO, "\x1b[32mINFO\033[1;0m")
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, datefmt=LOG_DATE)
log = logging.getLogger(__name__)


# ~ You can get custom salt
# ~ codecs.encode(os.urandom(16), 'hex')
# ~ but, not modify this file, modify in import file
SALT = b"c9548699d4e432dfd2b46adddafbb06d"

TIMEOUT = 10
LOG_NAME = "ZAZ"
FILE_NAME_CONFIG = "zaz-{}.json"

LEFT = 0
CENTER = 1
RIGHT = 2

CALC = "calc"
WRITER = "writer"
DRAW = "draw"
IMPRESS = "impress"
BASE = "base"
MATH = "math"
BASIC = "basic"
MAIN = "main"
TYPE_DOC = {
    CALC: "com.sun.star.sheet.SpreadsheetDocument",
    WRITER: "com.sun.star.text.TextDocument",
    DRAW: "com.sun.star.drawing.DrawingDocument",
    IMPRESS: "com.sun.star.presentation.PresentationDocument",
    BASE: "com.sun.star.sdb.DocumentDataSource",
    MATH: "com.sun.star.formula.FormulaProperties",
    BASIC: "com.sun.star.script.BasicIDE",
    MAIN: "com.sun.star.frame.StartModule",
}

OBJ_CELL = "ScCellObj"
OBJ_RANGE = "ScCellRangeObj"
OBJ_RANGES = "ScCellRangesObj"
TYPE_RANGES = (OBJ_CELL, OBJ_RANGE)

OBJ_SHAPE = "com.sun.star.comp.sc.ScShapeObj"
OBJ_SHAPES = "com.sun.star.drawing.SvxShapeCollection"
OBJ_GRAPHIC = "SwXTextGraphicObject"

OBJ_TEXTS = "SwXTextRanges"
OBJ_TEXT = "SwXTextRange"

CLSID = {
    "FORMULA": "078B7ABA-54FC-457F-8551-6147e776a997",
}

SERVICES = {
    "TEXT_EMBEDDED": "com.sun.star.text.TextEmbeddedObject",
    "TEXT_TABLE": "com.sun.star.text.TextTable",
    "GRAPHIC": "com.sun.star.text.GraphicObject",
}


# ~ from com.sun.star.sheet.FilterOperator import EMPTY, NO_EMPTY, EQUAL, NOT_EQUAL
class FilterOperator(IntEnum):
    EMPTY = 0
    NO_EMPTY = 1
    EQUAL = 2
    NOT_EQUAL = 3


# ~ https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlEditModel.html#a54d3ff280d892218d71e667f81ce99d4
class Border(IntEnum):
    NO_BORDER = 0
    BORDER = 1
    SIMPLE = 2


# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet.html#aa5aa6dbecaeb5e18a476b0a58279c57a
class ValidationType:
    from com.sun.star.sheet.ValidationType import ANY, WHOLE, DECIMAL, DATE, TIME, TEXT_LEN, LIST, CUSTOM


VT = ValidationType


# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet.html#aecf58149730f4c8c5c18c70f3c7c5db7
class ValidationAlertStyle:
    from com.sun.star.sheet.ValidationAlertStyle import STOP, WARNING, INFO, MACRO


VAS = ValidationAlertStyle


# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1ConditionOperator2.html
class ConditionOperator:
    from com.sun.star.sheet.ConditionOperator2 import (
        NONE,
        EQUAL,
        NOT_EQUAL,
        GREATER,
        GREATER_EQUAL,
        LESS,
        LESS_EQUAL,
        BETWEEN,
        NOT_BETWEEN,
        FORMULA,
        DUPLICATE,
        NOT_DUPLICATE,
    )


CO = ConditionOperator


class DataPilotFieldOrientation:
    from com.sun.star.sheet.DataPilotFieldOrientation import HIDDEN, COLUMN, ROW, PAGE, DATA


DPFO = DataPilotFieldOrientation


class CellInsertMode:
    from com.sun.star.sheet.CellInsertMode import DOWN, RIGHT, ROWS, COLUMNS


CIM = CellInsertMode


class CellDeleteMode:
    from com.sun.star.sheet.CellDeleteMode import UP, LEFT, ROWS, COLUMNS


CDM = CellDeleteMode


class FormButtonType:
    from com.sun.star.form.FormButtonType import PUSH, SUBMIT, RESET, URL


FBT = FormButtonType


class TextContentAnchorType:
    from com.sun.star.text.TextContentAnchorType import AT_PARAGRAPH, AS_CHARACTER, AT_PAGE, AT_FRAME, AT_CHARACTER


TCAT = TextContentAnchorType


OS = platform.system()
IS_WIN = OS == "Windows"
IS_MAC = OS == "Darwin"
USER = getpass.getuser()
PC = platform.node()
DESKTOP = os.environ.get("DESKTOP_SESSION", "")
INFO_DEBUG = f"{sys.version}\n\n{platform.platform()}\n\n" + "\n".join(sys.path)

PYTHON = "python"
if IS_WIN:
    PYTHON = "python.exe"

_MACROS = {}
_start = 0

SECONDS_DAY = 60 * 60 * 24
DIR = {
    "images": "images",
    "locales": "locales",
}

KEY = {
    "enter": 1280,
}

MODIFIERS = {
    "shift": KeyModifier.SHIFT,
    "ctrl": KeyModifier.MOD1,
    "alt": KeyModifier.MOD2,
    "ctrlmac": KeyModifier.MOD3,
}

# ~ Menus
NODE_MENUBAR = "private:resource/menubar/menubar"
MENUS = {
    "file": ".uno:PickList",
    "tools": ".uno:ToolsMenu",
    "help": ".uno:HelpMenu",
    "windows": ".uno:WindowList",
    "edit": ".uno:EditMenu",
    "view": ".uno:ViewMenu",
    "insert": ".uno:InsertMenu",
    "format": ".uno:FormatMenu",
    "styles": ".uno:FormatStylesMenu",
    "sheet": ".uno:SheetMenu",
    "data": ".uno:DataMenu",
    "table": ".uno:TableMenu",
    "form": ".uno:FormatFormMenu",
    "page": ".uno:PageMenu",
    "shape": ".uno:ShapeMenu",
    "slide": ".uno:SlideMenu",
    "show": ".uno:SlideShowMenu",
}

DEFAULT_MIME_TYPE = "png"
MIME_TYPE = {
    "png": "image/png",
    "jpg": "image/jpeg",
}

MESSAGES = {
    "es": {
        "OK": "Aceptar",
        "Cancel": "Cancelar",
        "Select path": "Seleccionar ruta",
        "Select directory": "Seleccionar directorio",
        "Select file": "Seleccionar archivo",
        "Incorrect user or password": "Nombre de usuario o contraseña inválidos",
        "Allow less secure apps in GMail": "Activa: Permitir aplicaciones menos segura en GMail",
    }
}


CTX = uno.getComponentContext()
SM = CTX.getServiceManager()


def create_instance(name: str, with_context: bool = False, args: Any = None) -> Any:
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    elif args:
        instance = SM.createInstanceWithArguments(name, (args,))
    else:
        instance = SM.createInstance(name)
    return instance


def get_app_config(node_name: str, key: str = ""):
    name = "com.sun.star.configuration.ConfigurationProvider"
    service = "com.sun.star.configuration.ConfigurationAccess"
    cp = create_instance(name, True)
    node = PropertyValue(Name="nodepath", Value=node_name)
    try:
        ca = cp.createInstanceWithArguments(service, (node,))
        if ca and not key:
            return ca
        if ca and ca.hasByName(key):
            return ca.getPropertyValue(key)
    except Exception as e:
        error(e)
        return ""


LANGUAGE = get_app_config("org.openoffice.Setup/L10N/", "ooLocale")
LANG = LANGUAGE.split("-")[0]
try:
    COUNTRY = LANGUAGE.split("-")[1]
except:
    COUNTRY = ""
LOCALE = Locale(LANG, COUNTRY, "")
NAME = TITLE = get_app_config("org.openoffice.Setup/Product", "ooName")
VERSION = get_app_config("org.openoffice.Setup/Product", "ooSetupVersion")

INFO_DEBUG = f"{NAME} v{VERSION} {LANGUAGE}\n\n{INFO_DEBUG}"

node = "/org.openoffice.Office.Calc/Calculate/Other/Date"
y = get_app_config(node, "YY")
m = get_app_config(node, "MM")
d = get_app_config(node, "DD")
DATE_OFFSET = datetime.date(y, m, d).toordinal()


def error(info):
    log.error(info)
    return


def debug(*args):
    data = [str(a) for a in args]
    log.debug("\t".join(data))
    return


def info(*args):
    data = [str(a) for a in args]
    log.info("\t".join(data))
    return


def save_log(path: str, data):
    with open(path, "a") as f:
        f.write(f"{str(now())[:19]} -{LOG_NAME}- ")
        pprint(data, stream=f)
    return


def catch_exception(f):
    @wraps(f)
    def func(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except Exception as e:
            name = f.__name__
            if IS_WIN:
                msgbox(traceback.format_exc())
            log.error(name, exc_info=True)

    return func


def inspect(obj: Any) -> None:
    zaz = create_instance("net.elmau.zaz.inspect")
    if hasattr(obj, "obj"):
        obj = obj.obj
    zaz.inspect(obj)
    return


def mri(obj: Any) -> None:
    m = create_instance("mytools.Mri")
    if m is None:
        msg = "Extension MRI not found"
        error(msg)
        return

    if hasattr(obj, "obj"):
        obj = obj.obj
    m.inspect(obj)
    return


def run_in_thread(fn):
    def run(*k, **kw):
        t = threading.Thread(target=fn, args=k, kwargs=kw)
        t.start()
        return t

    return run


def now(only_time: bool = False):
    now = datetime.datetime.now()
    if only_time:
        now = now.time()
    return now


def today():
    return datetime.date.today()


def _(msg):
    if LANG == "en":
        return msg

    if not LANG in MESSAGES:
        return msg

    return MESSAGES[LANG][msg]


def msgbox(message, title=TITLE, buttons=MSG_BUTTONS.BUTTONS_OK, type_msg="infobox"):
    """Create message box
    type_msg: infobox, warningbox, errorbox, querybox, messbox
    http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance("com.sun.star.awt.Toolkit")
    parent = toolkit.getDesktopWindow()
    box = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return box.execute()


def question(message, title=TITLE):
    result = msgbox(message, title, MSG_BUTTONS.BUTTONS_YES_NO, "querybox")
    return result == YES


def warning(message, title=TITLE):
    return msgbox(message, title, type_msg="warningbox")


def errorbox(message, title=TITLE):
    return msgbox(message, title, type_msg="errorbox")


def get_type_doc(obj: Any) -> str:
    for k, v in TYPE_DOC.items():
        if obj.supportsService(v):
            return k
    return ""


def _get_class_doc(obj: Any) -> Any:
    classes = {
        CALC: LOCalc,
        WRITER: LOWriter,
        DRAW: LODraw,
        IMPRESS: LOImpress,
        BASE: LOBase,
        MATH: LOMath,
        BASIC: LOBasic,
    }
    type_doc = get_type_doc(obj)
    return classes[type_doc](obj)


def dict_to_property(values: dict, uno_any: bool = False):
    ps = tuple([PropertyValue(Name=n, Value=v) for n, v in values.items()])
    if uno_any:
        ps = uno.Any("[]com.sun.star.beans.PropertyValue", ps)
    return ps


def _array_to_dict(values):
    d = {v[0]: v[1] for v in values}
    return d


def _property_to_dict(values):
    d = {v.Name: v.Value for v in values}
    return d


def json_dumps(data):
    return json.dumps(data, indent=4, sort_keys=True)


def json_loads(data):
    return json.loads(data)


def data_to_dict(data):
    if isinstance(data, (tuple, list)) and isinstance(data[0], (tuple, list)):
        return _array_to_dict(data)

    if isinstance(data, (tuple, list)) and isinstance(data[0], (PropertyValue, NamedValue)):
        return _property_to_dict(data)
    return {}


def _get_dispatch() -> Any:
    return create_instance("com.sun.star.frame.DispatchHelper")


# ~ https://wiki.documentfoundation.org/Development/DispatchCommands
# ~ Used only if not exists in API
def call_dispatch(frame: Any, url: str, args: dict = {}) -> None:
    dispatch = _get_dispatch()
    if hasattr(frame, "frame"):
        frame = frame.frame
    opt = dict_to_property(args)
    dispatch.executeDispatch(frame, url, "", 0, opt)
    return


def get_desktop():
    return create_instance("com.sun.star.frame.Desktop", True)


def _date_to_struct(value):
    if isinstance(value, datetime.datetime):
        d = DateTime()
        d.Year = value.year
        d.Month = value.month
        d.Day = value.day
        d.Hours = value.hour
        d.Minutes = value.minute
        d.Seconds = value.second
    elif isinstance(value, datetime.date):
        d = Date()
        d.Day = value.day
        d.Month = value.month
        d.Year = value.year
    elif isinstance(value, datetime.time):
        d = Time()
        d.Hours = value.hour
        d.Minutes = value.minute
        d.Seconds = value.second
    return d


def _struct_to_date(value):
    d = None
    if isinstance(value, Time):
        d = datetime.time(value.Hours, value.Minutes, value.Seconds)
    elif isinstance(value, Date):
        if value != Date():
            d = datetime.date(value.Year, value.Month, value.Day)
    elif isinstance(value, DateTime):
        if value.Year > 0:
            d = datetime.datetime(value.Year, value.Month, value.Day, value.Hours, value.Minutes, value.Seconds)
    return d


def _get_url_script(args: dict):
    library = args["library"]
    name = args["name"]
    language = args.get("language", "Python")
    location = args.get("location", "user")
    module = args.get("module", ".")

    if language == "Python":
        module = ".py$"
    elif language == "Basic":
        module = f".{module}."
        if location == "user":
            location = "application"

    url = "vnd.sun.star.script"
    url = f"{url}:{library}{module}{name}?language={language}&location={location}"
    return url


def _call_macro(args: dict):
    # ~ https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/Scripting_Framework_URI_Specification

    url = _get_url_script(args)
    args = args.get("args", ())

    service = "com.sun.star.script.provider.MasterScriptProviderFactory"
    factory = create_instance(service)
    script = factory.createScriptProvider("").getScript(url)
    result = script.invoke(args, None, None)[0]

    return result


def call_macro(args, in_thread=False):
    result = None
    if in_thread:
        t = threading.Thread(target=_call_macro, args=(args,))
        t.start()
    else:
        result = _call_macro(args)
    return result


def run(command, capture=False, split=False):
    if split:
        cmd = shlex.split(command)
        result = subprocess.run(cmd, capture_output=capture, text=True, shell=IS_WIN)
        if capture:
            result = result.stdout
        else:
            result = result.returncode
    else:
        if capture:
            result = subprocess.check_output(command, shell=True).decode()
        else:
            result = subprocess.Popen(command)
    return result


def popen(command):
    try:
        proc = subprocess.Popen(shlex.split(command), shell=IS_WIN, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        for line in proc.stdout:
            yield line.decode().rstrip()
    except Exception as e:
        error(e)
        yield (e.errno, e.strerror)


def sleep(seconds):
    time.sleep(seconds)
    return


class TimerThread(threading.Thread):
    def __init__(self, event, seconds, macro):
        threading.Thread.__init__(self)
        self.stopped = event
        self.seconds = seconds
        self.macro = macro

    def run(self):
        info("Timer started... {}".format(self.macro["name"]))
        while not self.stopped.wait(self.seconds):
            _call_macro(self.macro)
        info("Timer stopped... {}".format(self.macro["name"]))
        return


def start_timer(name, seconds, macro):
    global _MACROS
    _MACROS[name] = threading.Event()
    thread = TimerThread(_MACROS[name], seconds, macro)
    thread.start()
    return


def stop_timer(name):
    global _MACROS
    _MACROS[name].set()
    del _MACROS[name]
    return


def install_locales(path: str, domain: str = "base", dir_locales=DIR["locales"]):
    path_locales = _P.join(_P(path).path, dir_locales)
    try:
        lang = gettext.translation(domain, path_locales, languages=[LANG])
        lang.install()
        _ = lang.gettext
    except Exception as e:
        from gettext import gettext as _

        error(e)
    return _


def _export_image(obj, args):
    name = "com.sun.star.drawing.GraphicExportFilter"
    exporter = create_instance(name)
    path = _P.to_system(args["URL"])
    args = dict_to_property(args)
    exporter.setSourceDocument(obj)
    exporter.filter(args)
    return _P.exists(path)


def sha256(data):
    result = hashlib.sha256(data.encode()).hexdigest()
    return result


def sha512(data):
    result = hashlib.sha512(data.encode()).hexdigest()
    return result


def get_config(key="", prefix="conf", default={}):
    name_file = FILE_NAME_CONFIG.format(prefix)
    values = None
    # ~ path = _P.join(_P.config('UserConfig'), name_file)
    path = _P.join(_P.user_config, name_file)
    if not _P.exists(path):
        return default

    values = _P.from_json(path)
    if key:
        values = values.get(key, default)

    return values


def set_config(key, value, prefix="conf"):
    name_file = FILE_NAME_CONFIG.format(prefix)
    # ~ path = _P.join(_P.config('UserConfig'), name_file)
    path = _P.join(_P.user_config, name_file)
    values = get_config(default={}, prefix=prefix)
    values[key] = value
    result = _P.to_json(path, values)
    return result


def start():
    global _start

    _start = now()
    info(_start)
    return


def end(get_seconds: bool = False):
    global _start

    e = now()
    td = e - _start
    result = str(td)
    if get_seconds:
        result = td.total_seconds()
    return result


def get_epoch():
    n = now()
    return int(time.mktime(n.timetuple()))


def render(template, data):
    s = Template(template)
    return s.safe_substitute(**data)


def get_size_screen():
    res = ""
    if IS_WIN:
        user32 = ctypes.windll.user32
        res = f"{user32.GetSystemMetrics(0)}x{user32.GetSystemMetrics(1)}"
    else:
        try:
            args = 'xrandr | grep "*" | cut -d " " -f4'
            res = run(args, split=False)
        except Exception as e:
            error(e)
    return res.strip()


def url_open(url, data=None, headers={}, verify=True, get_json=False, timeout=TIMEOUT):
    err = ""
    req = Request(url)
    for k, v in headers.items():
        req.add_header(k, v)
    try:
        # ~ debug(url)
        if verify:
            if not data is None and isinstance(data, str):
                data = data.encode()
            response = urlopen(req, data=data, timeout=timeout)
        else:
            context = ssl._create_unverified_context()
            response = urlopen(req, data=data, timeout=timeout, context=context)
    except HTTPError as e:
        error(e)
        err = str(e)
    except URLError as e:
        error(e.reason)
        err = str(e.reason)
    except timeout:
        err = "timeout"
        error(err)
    else:
        headers = dict(response.info())
        result = response.read().decode()
        if get_json:
            result = json.loads(result)

    return result, headers, err


def _get_key(password):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=SALT, iterations=100000)
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    return key


def encrypt(data, password):
    from cryptography.fernet import Fernet

    f = Fernet(_get_key(password))
    if isinstance(data, str):
        data = data.encode()
    token = f.encrypt(data).decode()
    return token


def decrypt(token, password):
    from cryptography.fernet import Fernet, InvalidToken

    data = ""
    f = Fernet(_get_key(password))
    try:
        data = f.decrypt(token.encode()).decode()
    except InvalidToken as e:
        error("Invalid Token")
    return data


def switch_design_mode(doc):
    call_dispatch(doc.frame, ".uno:SwitchControlDesignMode")
    return


class SmtpServer(object):
    def __init__(self, config):
        self._server = None
        self._error = ""
        self._sender = ""
        self._is_connect = self._login(config)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    @property
    def is_connect(self):
        return self._is_connect

    @property
    def error(self):
        return self._error

    def _login(self, config):
        name = config["server"]
        port = config["port"]
        is_ssl = config["ssl"]
        self._sender = config["user"]
        hosts = "gmail" in name or "outlook" in name
        try:
            if is_ssl and hosts:
                self._server = smtplib.SMTP(name, port, timeout=TIMEOUT)
                self._server.ehlo()
                self._server.starttls()
                self._server.ehlo()
            elif is_ssl:
                self._server = smtplib.SMTP_SSL(name, port, timeout=TIMEOUT)
                self._server.ehlo()
            else:
                self._server = smtplib.SMTP(name, port, timeout=TIMEOUT)

            self._server.login(self._sender, config["password"])
            msg = "Connect to: {}".format(name)
            debug(msg)
            return True
        except smtplib.SMTPAuthenticationError as e:
            if "535" in str(e):
                self._error = _("Incorrect user or password")
                return False
            if "534" in str(e) and "gmail" in name:
                self._error = _("Allow less secure apps in GMail")
                return False
        except smtplib.SMTPException as e:
            self._error = str(e)
            return False
        except Exception as e:
            self._error = str(e)
            return False
        return False

    def _body(self, msg):
        body = msg.replace("\n", "<BR>")
        return body

    def send(self, message):
        # ~ file_name = 'attachment; filename={}'
        email = MIMEMultipart()
        email["From"] = self._sender
        email["To"] = message["to"]
        email["Cc"] = message.get("cc", "")
        email["Subject"] = message["subject"]
        email["Date"] = formatdate(localtime=True)
        if message.get("confirm", False):
            email["Disposition-Notification-To"] = email["From"]
        email.attach(MIMEText(self._body(message["body"]), "html"))

        paths = message.get("files", ())
        if isinstance(paths, str):
            paths = (paths,)
        for path in paths:
            fn = _P(path).file_name
            print("NAME", fn)
            part = MIMEBase("application", "octet-stream")
            part.set_payload(_P.read_bin(path))
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{fn}"')
            email.attach(part)

        receivers = email["To"].split(",") + email["CC"].split(",") + message.get("bcc", "").split(",")
        try:
            self._server.sendmail(self._sender, receivers, email.as_string())
            msg = "Email sent..."
            debug(msg)
            if message.get("path", ""):
                self.save_message(email, message["path"])
            return True
        except Exception as e:
            self._error = str(e)
            return False
        return False

    def save_message(self, email, path):
        mbox = mailbox.mbox(path, create=True)
        mbox.lock()
        try:
            msg = mailbox.mboxMessage(email)
            mbox.add(msg)
            mbox.flush()
        finally:
            mbox.unlock()
        return

    def close(self):
        try:
            self._server.quit()
            msg = "Close connection..."
            debug(msg)
        except:
            pass
        return


def _send_email(server, messages):
    with SmtpServer(server) as server:
        if server.is_connect:
            for msg in messages:
                server.send(msg)
        else:
            error(server.error)
    return server.error


def send_email(server, message):
    messages = message
    if isinstance(message, dict):
        messages = (message,)
    t = threading.Thread(target=_send_email, args=(server, messages))
    t.start()
    return


class ImapServer(object):
    def __init__(self, config):
        self._server = None
        self._error = ""
        self._is_connect = self._login(config)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    @property
    def is_connect(self):
        return self._is_connect

    @property
    def error(self):
        return self._error

    def _login(self, config):
        try:
            # ~ hosts = 'gmail' in config['server']
            if config["ssl"]:
                self._server = imaplib.IMAP4_SSL(config["server"], config["port"])
            else:
                self._server = imaplib.IMAP4(config["server"], config["port"])
            self._server.login(config["user"], config["password"])
            self._server.select()
            return True
        except imaplib.IMAP4.error as e:
            self._error = str(e)
            return False
        except Exception as e:
            self._error = str(e)
            return False
        return False

    def get_folders(self, exclude=()):
        folders = {}
        result, subdir = self._server.list()
        for s in subdir:
            print(s.decode("utf-8"))
        return folders

    def close(self):
        try:
            self._server.close()
            self._server.logout()
            msg = "Close connection..."
            debug(msg)
        except:
            pass
        return


# ~ Classes


class LOBaseObject(object):
    def __init__(self, obj):
        self._obj = obj

    def __setattr__(self, name, value):
        exists = hasattr(self, name)
        if not exists and not name in ("_obj", "_index", "_view"):
            setattr(self._obj, name, value)
        else:
            super().__setattr__(name, value)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj


class LODocument(object):
    def __init__(self, obj):
        self._obj = obj
        self._cc = self.obj.getCurrentController()
        self._undo = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    @property
    def obj(self):
        return self._obj

    @property
    def title(self):
        return self.obj.getTitle()

    @title.setter
    def title(self, value):
        self.obj.setTitle(value)

    @property
    def type(self):
        return self._type

    @property
    def uid(self):
        return self.obj.RuntimeUID

    @property
    def frame(self):
        return self._cc.getFrame()

    @property
    def is_saved(self):
        return self.obj.hasLocation()

    @property
    def is_modified(self):
        return self.obj.isModified()

    @property
    def is_read_only(self):
        return self.obj.isReadonly()

    @property
    def path(self):
        return _P.to_system(self.obj.URL)

    @property
    def dir(self):
        return _P(self.path).path

    @property
    def file_name(self):
        return _P(self.path).file_name

    @property
    def name(self):
        return _P(self.path).name

    @property
    def visible(self):
        w = self.frame.ContainerWindow
        return w.isVisible()

    @visible.setter
    def visible(self, value):
        w = self.frame.ContainerWindow
        w.setVisible(value)

    @property
    def zoom(self):
        return self._cc.ZoomValue

    @zoom.setter
    def zoom(self, value):
        self._cc.ZoomValue = value

    @property
    def undo(self):
        return self._undo

    @undo.setter
    def undo(self, value):
        self._undo = value
        um = self.obj.UndoManager
        if value:
            try:
                um.leaveUndoContext()
            except:
                pass
        else:
            um.enterHiddenUndoContext()

    def clear_undo(self):
        self.obj.getUndoManager().clear()
        return

    @property
    def selection(self):
        sel = self.obj.CurrentSelection
        return sel

    @property
    def table_auto_formats(self):
        taf = create_instance("com.sun.star.sheet.TableAutoFormats")
        return taf.ElementNames

    @property
    def status_bar(self):
        bar = self._cc.getStatusIndicator()
        return bar

    def create_instance(self, name):
        obj = self.obj.createInstance(name)
        return obj

    def set_focus(self):
        w = self.frame.ComponentWindow
        w.setFocus()
        return

    def copy(self):
        call_dispatch(self.frame, ".uno:Copy")
        return

    def insert_contents(self, args={}):
        call_dispatch(self.frame, ".uno:InsertContents", args)
        return

    def paste(self):
        sc = create_instance("com.sun.star.datatransfer.clipboard.SystemClipboard")
        transferable = sc.getContents()
        self._cc.insertTransferable(transferable)
        # ~ return self.obj.getCurrentSelection()
        return

    # ~ def select(self, obj):
    # ~ self._cc.select(obj)
    # ~ return

    def to_pdf(self, path: str = "", options: dict = {}):
        """
        https://wiki.documentfoundation.org/Macros/Python_Guide/PDF_export_filter_data
        """
        args = options.copy()
        stream = None
        path_pdf = "private:stream"
        if path:
            path_pdf = _P.to_url(path)

        filter_name = "{}_pdf_Export".format(self.type)
        filter_data = dict_to_property(args, True)
        args = {
            "FilterName": filter_name,
            "FilterData": filter_data,
        }
        if not path:
            stream = IOStream.output()
            args["OutputStream"] = stream

        opt = dict_to_property(args)
        try:
            self.obj.storeToURL(path_pdf, opt)
        except Exception as e:
            error(e)

        if not stream is None:
            stream = stream.buffer

        return stream

    def export(self, path: str = "", filter_name: str = "", options: dict = {}):
        FILTERS = {
            "xlsx": "Calc MS Excel 2007 XML",
            "xls": "MS Excel 97",
            "docx": "MS Word 2007 XML",
            "doc": "MS Word 97",
            "rtf": "Rich Text Format",
        }
        args = options.copy()
        stream = None
        path_target = "private:stream"
        if path:
            path_target = _P.to_url(path)

        filter_name = FILTERS.get(filter_name, filter_name)
        filter_data = dict_to_property(args, True)
        args = {
            "FilterName": filter_name,
            "FilterData": filter_data,
        }
        if not path:
            stream = IOStream.output()
            args["OutputStream"] = stream

        opt = dict_to_property(args)
        try:
            self.obj.storeToURL(path_target, opt)
        except Exception as e:
            error(e)

        if not stream is None:
            stream = stream.buffer

        return stream

    def save(self, path: str = "", options: dict = {}):
        if not path:
            self.obj.store()
            return

        args = options.copy()
        path_target = _P.to_url(path)

        opt = dict_to_property(args)
        try:
            self.obj.storeAsURL(path_target, opt)
        except Exception as e:
            error(e)

        return

    def close(self):
        self.obj.close(True)
        return


class LOCellStyle(LOBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def name(self):
        return self.obj.Name

    @property
    def properties(self):
        properties = self.obj.PropertySetInfo.Properties
        data = {p.Name: getattr(self.obj, p.Name) for p in properties}
        return data

    @properties.setter
    def properties(self, values):
        _set_properties(self.obj, values)


class LOCellStyles(object):
    def __init__(self, obj, doc):
        self._obj = obj
        self._doc = doc

    def __len__(self):
        return len(self.obj)

    def __getitem__(self, index):
        return LOCellStyle(self.obj[index])

    def __setitem__(self, key, value):
        self.obj[key] = value

    def __delitem__(self, key):
        if not isinstance(key, str):
            key = key.Name
        del self.obj[key]

    def __contains__(self, item):
        return item in self.obj

    @property
    def obj(self):
        return self._obj

    @property
    def names(self):
        return self.obj.ElementNames

    def new(self, name: str = ""):
        obj = self._doc.create_instance("com.sun.star.style.CellStyle")
        if name:
            self.obj[name] = obj
            obj = LOCellStyle(obj)
        return obj


class LOCalc(LODocument):
    def __init__(self, obj):
        super().__init__(obj)
        self._type = CALC
        self._sheets = obj.Sheets

    def __getitem__(self, index):
        return LOCalcSheet(self._sheets[index])

    def __setitem__(self, key, value):
        self._sheets[key] = value

    def __len__(self):
        return self._sheets.Count

    def __contains__(self, item):
        return item in self._sheets

    @property
    def names(self):
        names = self.obj.Sheets.ElementNames
        return names

    @property
    def selection(self):
        sel = self.obj.CurrentSelection
        if sel.ImplementationName in TYPE_RANGES:
            sel = LOCalcRange(sel)
        elif sel.ImplementationName in OBJ_RANGES:
            sel = LOCalcRanges(sel)
        elif sel.ImplementationName == OBJ_SHAPES:
            if len(sel) == 1:
                sel = LOShape(sel[0])
        else:
            debug(sel.ImplementationName)
        return sel

    @property
    def active(self):
        return LOCalcSheet(self._cc.ActiveSheet)

    @property
    def headers(self):
        return self._cc.ColumnRowHeaders

    @headers.setter
    def headers(self, value):
        self._cc.ColumnRowHeaders = value

    @property
    def tabs(self):
        return self._cc.SheetTabs

    @tabs.setter
    def tabs(self, value):
        self._cc.SheetTabs = value

    @property
    def cs(self):
        return self.cell_styles

    @property
    def cell_styles(self):
        obj = self.obj.StyleFamilies["CellStyles"]
        return LOCellStyles(obj, self)

    @property
    def db_ranges(self):
        # ~ return LOCalcDataBaseRanges(self.obj.DataBaseRanges)
        return self.obj.DatabaseRanges

    @property
    def ranges(self):
        obj = self.create_instance("com.sun.star.sheet.SheetCellRanges")
        return LOCalcRanges(obj)

    def get_ranges(self, address: str):
        ranges = self.ranges
        ranges.add([sheet[address] for sheet in self])
        return ranges

    def activate(self, sheet):
        obj = sheet
        if isinstance(sheet, LOCalcSheet):
            obj = sheet.obj
        elif isinstance(sheet, str):
            obj = self._sheets[sheet]
        self._cc.setActiveSheet(obj)
        return

    def new_sheet(self):
        s = self.create_instance("com.sun.star.sheet.Spreadsheet")
        return s

    def insert(self, name):
        names = name
        if isinstance(name, str):
            names = (name,)
        for n in names:
            self._sheets[n] = self.new_sheet()
        return LOCalcSheet(self._sheets[n])

    def move(self, name, pos=-1):
        index = pos
        if pos < 0:
            index = len(self)
        if isinstance(name, LOCalcSheet):
            name = name.name
        self._sheets.moveByName(name, index)
        return

    def remove(self, name):
        if isinstance(name, LOCalcSheet):
            name = name.name
        self._sheets.removeByName(name)
        return

    def copy_sheet(self, name, new_name="", pos=-1):
        if isinstance(name, LOCalcSheet):
            name = name.name
        index = pos
        if pos < 0:
            index = len(self)
        self._sheets.copyByName(name, new_name, index)
        return LOCalcSheet(self._sheets[new_name])

    def copy_from(self, doc, source="", target="", pos=-1):
        index = pos
        if pos < 0:
            index = len(self)

        names = source
        if not source:
            names = doc.names
        elif isinstance(source, str):
            names = (source,)

        new_names = target
        if not target:
            new_names = names
        elif isinstance(target, str):
            new_names = (target,)

        for i, name in enumerate(names):
            self._sheets.importSheet(doc.obj, name, index + i)
            self[index + i].name = new_names[i]

        return LOCalcSheet(self._sheets[index])

    def sort(self, reverse=False):
        names = sorted(self.names, reverse=reverse)
        for i, n in enumerate(names):
            self.move(n, i)
        return

    def render(self, data, sheet=None, clean=True):
        if sheet is None:
            sheet = self.active
        return sheet.render(data, clean=clean)

    def select(self, rango):
        self._cc.select(rango.obj)
        return


class LOChart(object):
    def __init__(self, name, obj, draw_page):
        self._name = name
        self._obj = obj
        self._eobj = self._obj.EmbeddedObject
        self._type = "Column"
        self._cell = None
        self._shape = self._get_shape(draw_page)
        self._pos = self._shape.Position

    def __getitem__(self, index):
        return LOBaseObject(self.diagram.getDataRowProperties(index))

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def name(self):
        return self._name

    @property
    def diagram(self):
        return self._eobj.Diagram

    @property
    def type(self):
        return self._type

    @type.setter
    def type(self, value):
        self._type = value
        if value == "Bar":
            self.diagram.Vertical = True
            return
        type_chart = f"com.sun.star.chart.{value}Diagram"
        self._eobj.setDiagram(self._eobj.createInstance(type_chart))

    @property
    def cell(self):
        return self._cell

    @cell.setter
    def cell(self, value):
        self._cell = value
        self._shape.Anchor = value.obj

    @property
    def position(self):
        return self._pos

    @position.setter
    def position(self, value):
        self._pos = value
        self._shape.Position = value

    def _get_shape(self, draw_page):
        for shape in draw_page:
            if shape.PersistName == self.name:
                break
        return shape


class LOSheetCharts(object):
    def __init__(self, obj, sheet):
        self._obj = obj
        self._sheet = sheet

    def __getitem__(self, index):
        return LOChart(index, self.obj[index], self._sheet.draw_page)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __contains__(self, item):
        return item in self.obj

    def __len__(self):
        return len(self.obj)

    @property
    def obj(self):
        return self._obj

    def new(self, name, pos_size, data):
        self.obj.addNewByName(name, pos_size, data, True, True)
        return LOChart(name, self.obj[name], self._sheet.draw_page)


class LOSheetTableField(object):
    def __init__(self, obj):
        self._obj = obj

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def name(self):
        return self.obj.Name

    @property
    def orientation(self):
        return self.obj.Orientation

    @orientation.setter
    def orientation(self, value):
        self.obj.Orientation = value


# ~ com.sun.star.sheet.DataPilotFieldOrientation.ROW
class LOSheetTable(object):
    def __init__(self, obj):
        self._obj = obj
        self._source = None

    def __getitem__(self, index):
        field = self.obj.DataPilotFields[index]
        return LOSheetTableField(field)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def filter(self):
        return self.obj.ShowFilterButton

    @filter.setter
    def filter(self, value):
        self.obj.ShowFilterButton = value

    @property
    def source(self):
        return self._source

    @source.setter
    def source(self, value):
        self._source = value
        self.obj.SourceRange = value.range_address

    @property
    def rows(self):
        return self.obj.RowFields

    @rows.setter
    def rows(self, values):
        if not isinstance(values, tuple):
            values = (values,)
        for v in values:
            with self[v] as f:
                f.orientation = DPFO.ROW

    @property
    def columns(self):
        return self.obj.ColumnFields

    @columns.setter
    def columns(self, values):
        if not isinstance(values, tuple):
            values = (values,)
        for v in values:
            with self[v] as f:
                f.orientation = DPFO.COLUMN

    @property
    def data(self):
        return self.obj.DataFields

    @data.setter
    def data(self, values):
        if not isinstance(values, tuple):
            values = (values,)
        for v in values:
            with self[v] as f:
                f.orientation = DPFO.DATA


class LOSheetTables(object):
    def __init__(self, obj, sheet):
        self._obj = obj
        self._sheet = sheet

    def __getitem__(self, index):
        return LOSheetTable(self.obj[index])

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __contains__(self, item):
        return item in self.obj

    @property
    def obj(self):
        return self._obj

    @property
    def count(self):
        return self.obj.Count

    @property
    def names(self):
        return self.obj.ElementNames

    def new(self, name, target):
        table = self.obj.createDataPilotDescriptor()
        self.obj.insertNewByName(name, target.address, table)
        return LOSheetTable(self.obj[name])

    def remove(self, name):
        self.obj.removeByName(name)
        return


# ~ class LOFormControl(LOBaseObject):
class LOFormControl:
    EVENTS = {
        "action": "actionPerformed",
        "click": "mousePressed",
    }
    TYPES = {
        "actionPerformed": "XActionListener",
        "mousePressed": "XMouseListener",
    }

    def __init__(self, obj, view, form):
        self._obj = obj
        self._view = view
        self._form = form
        self._m = view.Model
        self._index = -1

    # ~ def __setattr__(self, name, value):
    # ~ if name in ('_form', '_view', '_m', '_index'):
    # ~ self.__dict__[name] = value
    # ~ else:
    # ~ super().__setattr__(name, value)

    def __str__(self):
        return f"{self.name} ({self.type}) {[self.index]}"

    @property
    def obj(self):
        return self._obj

    @property
    def form(self):
        return self._form

    @property
    def doc(self):
        return self.obj.Parent.Forms.Parent

    @property
    def name(self):
        return self._m.Name

    @name.setter
    def name(self, value):
        self._m.Name = value

    @property
    def tag(self):
        return self._m.Tag

    @tag.setter
    def tag(self, value):
        self._m.Tag = value

    @property
    def index(self):
        return self._index

    @index.setter
    def index(self, value):
        self._index = value

    @property
    def enabled(self):
        return self._m.Enabled

    @enabled.setter
    def enabled(self, value):
        self._m.Enabled = value

    @property
    def anchor(self):
        return self.obj.Anchor

    @anchor.setter
    def anchor(self, value):
        size = None
        if hasattr(value, "obj"):
            size = getattr(value, "size", None)
            value = value.obj
        self.obj.Anchor = value
        if not size is None:
            self.size = size
            try:
                self.obj.ResizeWithCell = True
            except:
                pass

    @property
    def size(self):
        return self.obj.Size

    @size.setter
    def size(self, value):
        self.obj.Size = value

    @property
    def events(self):
        return self.form.getScriptEvents(self.index)

    def add_event(self, name, macro):
        if not "name" in macro:
            macro["name"] = "{}_{}".format(self.name, name)

        event = ScriptEventDescriptor()
        event.AddListenerParam = ""
        event.EventMethod = self.EVENTS[name]
        event.ListenerType = self.TYPES[event.EventMethod]
        event.ScriptCode = _get_url_script(macro)
        event.ScriptType = "Script"

        for ev in self.events:
            if ev.EventMethod == event.EventMethod and ev.ListenerType == event.ListenerType:
                self.form.revokeScriptEvent(self.index, event.ListenerType, event.EventMethod, event.AddListenerParam)
                break

        self.form.registerScriptEvent(self.index, event)
        return

    def set_focus(self):
        self._view.setFocus()
        return


class LOFormControlLabel(LOFormControl):
    def __init__(self, obj, view, form):
        super().__init__(obj, view, form)

    @property
    def type(self):
        return "label"

    @property
    def value(self):
        return self._m.Label

    @value.setter
    def value(self, value):
        self._m.Label = value


class LOFormControlText(LOFormControl):
    def __init__(self, obj, view, form):
        super().__init__(obj, view, form)

    @property
    def type(self):
        return "text"

    @property
    def value(self):
        return self._m.Text

    @value.setter
    def value(self, value):
        self._m.Text = value


class LOFormControlButton(LOFormControl):
    def __init__(self, obj, view, form):
        super().__init__(obj, view, form)

    @property
    def type(self):
        return "button"

    @property
    def value(self):
        return self._m.Label

    @value.setter
    def value(self, value):
        self._m.Text = Label

    @property
    def url(self):
        return self._m.TargetURL

    @url.setter
    def url(self, value):
        self._m.TargetURL = value
        self._m.ButtonType = FormButtonType.URL


FORM_CONTROL_CLASS = {
    "label": LOFormControlLabel,
    "text": LOFormControlText,
    "button": LOFormControlButton,
}


class LOForm(object):
    MODELS = {
        "label": "com.sun.star.form.component.FixedText",
        "text": "com.sun.star.form.component.TextField",
        "button": "com.sun.star.form.component.CommandButton",
    }

    def __init__(self, obj, draw_page):
        self._obj = obj
        self._dp = draw_page
        self._controls = {}
        self._init_controls()

    def __getitem__(self, index):
        control = self.obj[index]
        return self._controls[control.Name]

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __contains__(self, item):
        return item in self.obj

    def __len__(self):
        return len(self.obj)

    def __str__(self):
        return f"Form: {self.name}"

    def _init_controls(self):
        types = {
            "com.sun.star.form.OFixedTextModel": "label",
            "com.sun.star.form.OEditModel": "text",
            "com.sun.star.form.OButtonModel": "button",
        }
        for i, control in enumerate(self.obj):
            name = control.Name
            tipo = types[control.ImplementationName]
            view = self.doc.CurrentController.getControl(control)
            control = FORM_CONTROL_CLASS[tipo](control, view, self._obj)
            control.index = i
            setattr(self, name, control)
            self._controls[name] = control
        return

    @property
    def obj(self):
        return self._obj

    @property
    def name(self):
        return self.obj.Name

    @name.setter
    def name(self, value):
        self.obj.Name = value

    @property
    def source(self):
        return self.obj.DataSourceName

    @source.setter
    def source(self, value):
        self.obj.DataSourceName = value

    @property
    def type(self):
        return self.obj.CommandType

    @type.setter
    def type(self, value):
        self.obj.CommandType = value

    @property
    def command(self):
        return self.obj.Command

    @command.setter
    def command(self, value):
        self.obj.Command = value

    @property
    def doc(self):
        return self.obj.Parent.Parent

    def _special_properties(self, tipo, args):
        if tipo == "button":
            # ~ if 'ImageURL' in args:
            # ~ args['ImageURL'] = self._set_image_url(args['ImageURL'])
            args["FocusOnClick"] = args.get("FocusOnClick", False)
            return args
        return args

    def add(self, args):
        name = args["Name"]
        tipo = args.pop("Type").lower()
        w = args.pop("Width", 1000)
        h = args.pop("Height", 200)
        x = args.pop("X", 0)
        y = args.pop("Y", 0)
        control = self.doc.createInstance("com.sun.star.drawing.ControlShape")
        control.setSize(Size(w, h))
        control.setPosition(Point(x, y))
        model = self.doc.createInstance(self.MODELS[tipo])
        args = self._special_properties(tipo, args)
        _set_properties(model, args)
        control.Control = model
        index = len(self)
        self.obj.insertByIndex(index, model)
        self._dp.add(control)
        view = self.doc.CurrentController.getControl(self.obj.getByName(name))
        control = FORM_CONTROL_CLASS[tipo](control, view, self.obj)
        control.index = index
        setattr(self, name, control)
        self._controls[name] = control
        return control


class LOSheetForms(object):
    def __init__(self, draw_page):
        self._dp = draw_page
        self._obj = draw_page.Forms

    def __getitem__(self, index):
        return LOForm(self.obj[index], self._dp)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __contains__(self, item):
        return item in self.obj

    def __len__(self):
        return len(self.obj)

    @property
    def obj(self):
        return self._obj

    @property
    def doc(self):
        return self.obj.Parent

    @property
    def count(self):
        return len(self)

    @property
    def names(self):
        return self.obj.ElementNames

    def insert(self, name=""):
        if not name:
            name = f"form{self.count + 1}"
        form = self.doc.createInstance("com.sun.star.form.component.Form")
        self.obj.insertByName(name, form)
        return LOForm(form, self._dp)

    def remove(self, index):
        if isinstance(index, int):
            self.obj.removeByIndex(index)
        else:
            self.obj.removeByName(index)
        return


# ~ IsFiltered,
# ~ IsManualPageBreak,
# ~ IsStartOfNewPage
class LOSheetRows(object):
    def __init__(self, sheet, obj):
        self._sheet = sheet
        self._obj = obj

    def __getitem__(self, index):
        if isinstance(index, int):
            rows = LOSheetRows(self._sheet, self.obj[index])
        else:
            rango = self._sheet[index.start : index.stop, 0:]
            rows = LOSheetRows(self._sheet, rango.obj.Rows)
        return rows

    def __len__(self):
        return self.obj.Count

    @property
    def obj(self):
        return self._obj

    @property
    def visible(self):
        return self._obj.IsVisible

    @visible.setter
    def visible(self, value):
        self._obj.IsVisible = value

    @property
    def color(self):
        return self.obj.CellBackColor

    @color.setter
    def color(self, value):
        self.obj.CellBackColor = value

    @property
    def is_transparent(self):
        return self.obj.IsCellBackgroundTransparent

    @is_transparent.setter
    def is_transparent(self, value):
        self.obj.IsCellBackgroundTransparent = value

    @property
    def height(self):
        return self.obj.Height

    @height.setter
    def height(self, value):
        self.obj.Height = value

    def optimal(self):
        self.obj.OptimalHeight = True
        return

    def insert(self, index, count):
        self.obj.insertByIndex(index, count)
        return

    def remove(self, index, count):
        self.obj.removeByIndex(index, count)
        return


# ~ IsManualPageBreak,
# ~ IsStartOfNewPage
class LOSheetColumns(object):
    def __init__(self, sheet, obj):
        self._sheet = sheet
        self._obj = obj

    def __getitem__(self, index):
        if isinstance(index, (int, str)):
            rows = LOSheetColumns(self._sheet, self.obj[index])
        else:
            rango = self._sheet[0, index.start : index.stop]
            rows = LOSheetColumns(self._sheet, rango.obj.Columns)
        return rows

    def __len__(self):
        return self.obj.Count

    @property
    def obj(self):
        return self._obj

    @property
    def visible(self):
        return self._obj.IsVisible

    @visible.setter
    def visible(self, value):
        self._obj.IsVisible = value

    @property
    def width(self):
        return self.obj.Width

    @width.setter
    def width(self, value):
        self.obj.Width = value

    def optimal(self):
        self.obj.OptimalWidth = True
        return

    def insert(self, index, count):
        self.obj.insertByIndex(index, count)
        return

    def remove(self, index, count):
        self.obj.removeByIndex(index, count)
        return


class LOCalcSheet(object):
    def __init__(self, obj):
        self._obj = obj

    def __getitem__(self, index):
        return LOCalcRange(self.obj[index])

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __str__(self):
        return f"easymacro.LOCalcSheet: {self.name}"

    @property
    def obj(self):
        return self._obj

    @property
    def name(self):
        return self._obj.Name

    @name.setter
    def name(self, value):
        self._obj.Name = value

    @property
    def code_name(self):
        return self._obj.CodeName

    @code_name.setter
    def code_name(self, value):
        self._obj.CodeName = value

    @property
    def visible(self):
        return self._obj.IsVisible

    @visible.setter
    def visible(self, value):
        self._obj.IsVisible = value

    @property
    def is_protected(self):
        return self._obj.isProtected()

    @property
    def password(self):
        return ""

    @visible.setter
    def password(self, value):
        self.obj.protect(value)

    def unprotect(self, value):
        try:
            self.obj.unprotect(value)
            return True
        except:
            pass
        return False

    @property
    def color(self):
        return self._obj.TabColor

    @color.setter
    def color(self, value):
        self._obj.TabColor = get_color(value)

    @property
    def used_area(self):
        cursor = self.get_cursor()
        cursor.gotoEndOfUsedArea(True)
        return LOCalcRange(self[cursor.AbsoluteName].obj)

    @property
    def draw_page(self):
        return LODrawPage(self.obj.DrawPage)

    @property
    def dp(self):
        return self.draw_page

    @property
    def shapes(self):
        return self.draw_page

    @property
    def doc(self):
        return LOCalc(self.obj.DrawPage.Forms.Parent)

    @property
    def charts(self):
        return LOSheetCharts(self.obj.Charts, self)

    @property
    def tables(self):
        return LOSheetTables(self.obj.DataPilotTables, self)

    @property
    def rows(self):
        return LOSheetRows(self, self.obj.Rows)

    @property
    def columns(self):
        return LOSheetColumns(self, self.obj.Columns)

    @property
    def forms(self):
        return LOSheetForms(self.obj.DrawPage)

    @property
    def events(self):
        names = ("OnFocus", "OnUnfocus", "OnSelect", "OnDoubleClick", "OnRightClick", "OnChange", "OnCalculate")
        evs = self.obj.Events
        events = {n: _property_to_dict(evs.getByName(n)) for n in names if evs.getByName(n)}
        return events

    @events.setter
    def events(self, values):
        pv = "[]com.sun.star.beans.PropertyValue"
        ev = self.obj.Events
        for name, v in values.items():
            url = _get_url_script(v)
            args = dict_to_property(dict(EventType="Script", Script=url))
            # ~ e.replaceByName(k, args)
            uno.invoke(ev, "replaceByName", (name, uno.Any(pv, args)))

    @property
    def search_descriptor(self):
        return self.obj.createSearchDescriptor()

    @property
    def replace_descriptor(self):
        return self.obj.createReplaceDescriptor()

    def activate(self):
        self.doc.activate(self.obj)
        return

    # ~ ???
    def clean(self):
        doc = self.doc
        sheet = doc.create_instance("com.sun.star.sheet.Spreadsheet")
        doc._sheets.replaceByName(self.name, sheet)
        return

    def move(self, pos=-1):
        index = pos
        if pos < 0:
            index = len(self.doc)
        self.doc._sheets.moveByName(self.name, index)
        return

    def remove(self):
        self.doc._sheets.removeByName(self.name)
        return

    def copy(self, new_name="", pos=-1):
        index = pos
        if pos < 0:
            index = len(self.doc)
        self.doc._sheets.copyByName(self.name, new_name, index)
        return LOCalcSheet(self.doc._sheets[new_name])

    def copy_to(self, doc, target="", pos=-1):
        index = pos
        if pos < 0:
            index = len(doc)

        new_name = target or self.name
        new_pos = doc._sheets.importSheet(self.doc.obj, self.name, index)
        sheet = doc[new_pos]
        sheet.name = new_name
        return sheet

    def get_cursor(self, cell=None):
        if cell is None:
            cursor = self.obj.createCursor()
        else:
            cursor = self.obj.createCursorByRange(cell)
        return cursor

    def render(self, data, rango=None, clean=True):
        if rango is None:
            rango = self.used_area
        return rango.render(data, clean)

    def find(self, search_string, rango=None):
        if rango is None:
            rango = self.used_area
        return rango.find(search_string)


class LOCalcRange(object):
    def __init__(self, obj):
        self._obj = obj
        self._sd = None
        self._is_cell = obj.ImplementationName == OBJ_CELL

    def __getitem__(self, index):
        return LOCalcRange(self.obj[index])

    def __iter__(self):
        self._r = 0
        self._c = 0
        return self

    def __next__(self):
        try:
            rango = self[self._r, self._c]
        except Exception as e:
            raise StopIteration
        self._c += 1
        if self._c == self.columns:
            self._c = 0
            self._r += 1
        return rango

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __contains__(self, item):
        return item.in_range(self)

    def __str__(self):
        if self.is_none:
            s = "Range: None"
        else:
            s = f"Range: {self.name}"
        return s

    @property
    def obj(self):
        return self._obj

    @property
    def is_none(self):
        return self.obj is None

    @property
    def is_cell(self):
        return self._is_cell

    @property
    def back_color(self):
        return self._obj.CellBackColor

    @back_color.setter
    def back_color(self, value):
        self._obj.CellBackColor = get_color(value)

    @property
    def dp(self):
        return self.sheet.dp

    @property
    def sheet(self):
        return LOCalcSheet(self.obj.Spreadsheet)

    @property
    def doc(self):
        doc = self.obj.Spreadsheet.DrawPage.Forms.Parent
        return LODocument(doc)

    @property
    def name(self):
        return self.obj.AbsoluteName

    @property
    def code_name(self):
        name = self.name.replace("$", "").replace(".", "_").replace(":", "")
        return name

    @property
    def columns(self):
        return self.obj.Columns.Count

    @property
    def column(self):
        c1 = self.address.Column
        c2 = c1 + 1
        ra = self.current_region.range_address
        r1 = ra.StartRow
        r2 = ra.EndRow + 1
        return LOCalcRange(self.sheet[r1:r2, c1:c2].obj)

    @property
    def rows(self):
        return LOSheetRows(self.sheet, self.obj.Rows)

    @property
    def row(self):
        r1 = self.address.Row
        r2 = r1 + 1
        ra = self.current_region.range_address
        c1 = ra.StartColumn
        c2 = ra.EndColumn + 1
        return LOCalcRange(self.sheet[r1:r2, c1:c2].obj)

    @property
    def type(self):
        return self.obj.Type

    @property
    def error(self):
        return self.obj.getError()

    @property
    def value(self):
        v = None
        if self.type == VALUE:
            v = self.obj.getValue()
        elif self.type == TEXT:
            v = self.obj.getString()
        elif self.type == FORMULA:
            v = self.obj.getFormula()
        return v

    @value.setter
    def value(self, data):
        if isinstance(data, str):
            if data[0] in "=":
                self.obj.setFormula(data)
            else:
                self.obj.setString(data)
        elif isinstance(data, Decimal):
            self.obj.setValue(float(data))
        elif isinstance(data, (int, float, bool)):
            self.obj.setValue(data)
        elif isinstance(data, datetime.datetime):
            d = data.toordinal()
            t = (data - datetime.datetime.fromordinal(d)).seconds / SECONDS_DAY
            self.obj.setValue(d - DATE_OFFSET + t)
        elif isinstance(data, datetime.date):
            d = data.toordinal()
            self.obj.setValue(d - DATE_OFFSET)
        elif isinstance(data, datetime.time):
            d = (data.hour * 3600 + data.minute * 60 + data.second) / SECONDS_DAY
            self.obj.setValue(d)

    @property
    def date(self):
        value = int(self.obj.Value)
        date = datetime.date.fromordinal(value + DATE_OFFSET)
        return date

    @property
    def time(self):
        seconds = self.obj.Value * SECONDS_DAY
        time_delta = datetime.timedelta(seconds=seconds)
        time = (datetime.datetime.min + time_delta).time()
        return time

    @property
    def datetime(self):
        return datetime.datetime.combine(self.date, self.time)

    @property
    def data(self):
        return self.obj.getDataArray()

    @data.setter
    def data(self, values):
        if self._is_cell:
            self.to_size(len(values), len(values[0])).data = values
        else:
            self.obj.setDataArray(values)

    @property
    def dict(self):
        rows = self.data
        k = rows[0]
        data = [dict(zip(k, r)) for r in rows[1:]]
        return data

    @dict.setter
    def dict(self, values):
        data = [tuple(values[0].keys())]
        data += [tuple(d.values()) for d in values]
        self.data = data

    @property
    def formula(self):
        return self.obj.getFormulaArray()

    @formula.setter
    def formula(self, values):
        self.obj.setFormulaArray(values)

    @property
    def array_formula(self):
        return self.obj.ArrayFormula

    @array_formula.setter
    def array_formula(self, value):
        self.obj.ArrayFormula = value

    @property
    def address(self):
        return self.obj.CellAddress

    @property
    def range_address(self):
        return self.obj.RangeAddress

    @property
    def cursor(self):
        cursor = self.obj.Spreadsheet.createCursorByRange(self.obj)
        return cursor

    @property
    def current_region(self):
        cursor = self.cursor
        cursor.collapseToCurrentRegion()
        return LOCalcRange(self.sheet[cursor.AbsoluteName].obj)

    @property
    def next_cell(self):
        a = self.current_region.range_address
        col = a.StartColumn
        row = a.EndRow + 1
        return LOCalcRange(self.sheet[row, col].obj)

    @property
    def position(self):
        return self.obj.Position

    @property
    def size(self):
        return self.obj.Size

    @property
    def possize(self):
        data = {
            "Width": self.size.Width,
            "Height": self.size.Height,
            "X": self.position.X,
            "Y": self.position.Y,
        }
        return data

    @property
    def visible(self):
        cursor = self.cursor
        rangos = cursor.queryVisibleCells()
        rangos = LOCalcRanges(rangos)
        return rangos

    @property
    def merged_area(self):
        cursor = self.cursor
        cursor.collapseToMergedArea()
        rango = LOCalcRange(self.sheet[cursor.AbsoluteName].obj)
        return rango

    @property
    def empty(self):
        cursor = self.sheet.get_cursor(self.obj)
        cursor = self.cursor
        rangos = cursor.queryEmptyCells()
        rangos = [LOCalcRange(self.sheet[r.AbsoluteName].obj) for r in rangos]
        return tuple(rangos)

    @property
    def merge(self):
        return self.obj.IsMerged

    @merge.setter
    def merge(self, value):
        self.obj.merge(value)

    @property
    def style(self):
        return self.obj.CellStyle

    @style.setter
    def style(self, value):
        self.obj.CellStyle = value

    @property
    def auto_format(self):
        return ""

    @auto_format.setter
    def auto_format(self, value):
        self.obj.autoFormat(value)

    @property
    def validation(self):
        return self.obj.Validation

    @validation.setter
    def validation(self, values):
        current = self.validation
        if not values:
            current.Type = ValidationType.ANY
            current.ShowInputMessage = False
        else:
            is_list = False
            for k, v in values.items():
                if k == "Type" and v == VT.LIST:
                    is_list = True
                if k == "Formula1" and is_list:
                    if isinstance(v, (tuple, list)):
                        v = ";".join(['"{}"'.format(i) for i in v])
                setattr(current, k, v)
        self.obj.Validation = current

    def select(self):
        self.doc._cc.select(self.obj)
        return

    def search(self, options, find_all=True):
        rangos = None

        descriptor = self.sheet.search_descriptor
        descriptor.setSearchString(options["Search"])
        descriptor.SearchCaseSensitive = options.get("CaseSensitive", False)
        descriptor.SearchWords = options.get("Words", False)
        if hasattr(descriptor, "SearchRegularExpression"):
            descriptor.SearchRegularExpression = options.get("RegularExpression", False)
        if hasattr(descriptor, "SearchType") and "Type" in options:
            descriptor.SearchType = options["Type"]

        if find_all:
            found = self.obj.findAll(descriptor)
        else:
            found = self.obj.findFirst(descriptor)

        if found:
            if found.ImplementationName == OBJ_CELL:
                rangos = LOCalcRange(found)
            else:
                rangos = [LOCalcRange(f) for f in found]

        return rangos

    def replace(self, options):
        descriptor = self.sheet.replace_descriptor
        descriptor.setSearchString(options["Search"])
        descriptor.setReplaceString(options["Replace"])
        descriptor.SearchCaseSensitive = options.get("CaseSensitive", False)
        descriptor.SearchWords = options.get("Words", False)
        if hasattr(descriptor, "SearchRegularExpression"):
            descriptor.SearchRegularExpression = options.get("RegularExpression", False)
        if hasattr(descriptor, "SearchType") and "Type" in options:
            descriptor.SearchType = options["Type"]
        count = self.obj.replaceAll(descriptor)
        return count

    def in_range(self, rango):
        if isinstance(rango, LOCalcRange):
            address = rango.range_address
        else:
            address = rango.RangeAddress
        result = self.cursor.queryIntersection(address)
        return bool(result.Count)

    def offset(self, rows=0, cols=1):
        ra = self.range_address
        col = ra.EndColumn + cols
        row = ra.EndRow + rows
        return LOCalcRange(self.sheet[row, col].obj)

    def to_size(self, rows, cols):
        cursor = self.cursor
        cursor.collapseToSize(cols, rows)
        return LOCalcRange(self.sheet[cursor.AbsoluteName].obj)

    def move(self, target):
        sheet = self.sheet.obj
        sheet.moveRange(target.address, self.range_address)
        return

    def insert(self, insert_mode=CIM.DOWN):
        sheet = self.sheet.obj
        sheet.insertCells(self.range_address, insert_mode)
        return

    def delete(self, delete_mode=CDM.UP):
        sheet = self.sheet.obj
        sheet.removeRange(self.range_address, delete_mode)
        return

    def copy_from(self, source):
        self.sheet.obj.copyRange(self.address, source.range_address)
        return

    def copy_to(self, target):
        self.sheet.obj.copyRange(target.address, self.range_address)
        return

    # ~ def copy_to(self, cell, formula=False):
    # ~ rango = cell.to_size(self.rows, self.columns)
    # ~ if formula:
    # ~ rango.formula = self.formula
    # ~ else:
    # ~ rango.data = self.data
    # ~ return

    # ~ def copy_from(self, rango, formula=False):
    # ~ data = rango
    # ~ if isinstance(rango, LOCalcRange):
    # ~ if formula:
    # ~ data = rango.formula
    # ~ else:
    # ~ data = rango.data
    # ~ rows = len(data)
    # ~ cols = len(data[0])
    # ~ if formula:
    # ~ self.to_size(rows, cols).formula = data
    # ~ else:
    # ~ self.to_size(rows, cols).data = data
    # ~ return

    def optimal_width(self):
        self.obj.Columns.OptimalWidth = True
        return

    def clean_render(self, template="\{(\w.+)\}"):
        self._sd.SearchRegularExpression = True
        self._sd.setSearchString(template)
        self.obj.replaceAll(self._sd)
        return

    def render(self, data, clean=True):
        self._sd = self.sheet.obj.createSearchDescriptor()
        self._sd.SearchCaseSensitive = False
        for k, v in data.items():
            cell = self._render_value(k, v)
        return cell

    def _render_value(self, key, value, parent=""):
        cell = None
        if isinstance(value, dict):
            for k, v in value.items():
                # ~ print(1, 'RENDER', k, v)
                cell = self._render_value(k, v, key)
            return cell
        elif isinstance(value, (list, tuple)):
            self._render_list(key, value)
            return

        search = f"{{{key}}}"
        if parent:
            search = f"{{{parent}.{key}}}"
        ranges = self.find_all(search)

        if ranges is None:
            return

        # ~ for cell in ranges or range(0):
        for cell in ranges:
            self._set_new_value(cell, search, value)
        return LOCalcRange(cell)

    def _set_new_value(self, cell, search, value):
        if not cell.ImplementationName == "ScCellObj":
            return

        if isinstance(value, str):
            pattern = re.compile(search, re.IGNORECASE)
            new_value = pattern.sub(value, cell.String)
            cell.String = new_value
        else:
            LOCalcRange(cell).value = value
        return

    def _render_list(self, key, rows):
        for row in rows:
            for k, v in row.items():
                self._render_value(k, v)
        return

    def find(self, search_string):
        if self._sd is None:
            self._sd = self.sheet.obj.createSearchDescriptor()
            self._sd.SearchCaseSensitive = False

        self._sd.setSearchString(search_string)
        cell = self.obj.findFirst(self._sd)
        if cell:
            cell = LOCalcRange(cell)
        return cell

    def find_all(self, search_string):
        if self._sd is None:
            self._sd = self.sheet.obj.createSearchDescriptor()
            self._sd.SearchCaseSensitive = False

        self._sd.setSearchString(search_string)
        ranges = self.obj.findAll(self._sd)
        return ranges

    def filter(self, args, with_headers=True):
        ff = TableFilterField()
        ff.Field = args["Field"]
        ff.Operator = args["Operator"]
        if isinstance(args["Value"], str):
            ff.IsNumeric = False
            ff.StringValue = args["Value"]
        else:
            ff.IsNumeric = True
            ff.NumericValue = args["Value"]

        fd = self.obj.createFilterDescriptor(True)
        fd.ContainsHeader = with_headers
        fd.FilterFields = (ff,)
        # ~ self.obj.AutoFilter = True
        self.obj.filter(fd)
        return

    def copy_format_from(self, rango):
        rango.select()
        self.doc.copy()
        self.select()
        args = {
            "Flags": "T",
            "MoveMode": 4,
        }
        url = ".uno:InsertContents"
        call_dispatch(self.doc.frame, url, args)
        return

    def to_image(self):
        self.select()
        self.doc.copy()
        args = {"SelectedFormat": 141}
        url = ".uno:ClipboardFormatItems"
        call_dispatch(self.doc.frame, url, args)
        return self.sheet.shapes[-1]

    def insert_image(self, path, options={}):
        args = options.copy()
        ps = self.possize
        args["Width"] = args.get("Width", ps["Width"])
        args["Height"] = args.get("Height", ps["Height"])
        args["X"] = args.get("X", ps["X"])
        args["Y"] = args.get("Y", ps["Y"])
        # ~ img.ResizeWithCell = True
        img = self.sheet.dp.insert_image(path, args)
        img.anchor = self.obj
        args.clear()
        return img

    def insert_shape(self, tipo, args={}):
        ps = self.possize
        args["Width"] = args.get("Width", ps["Width"])
        args["Height"] = args.get("Height", ps["Height"])
        args["X"] = args.get("X", ps["X"])
        args["Y"] = args.get("Y", ps["Y"])

        shape = self.sheet.dp.add(tipo, args)
        shape.anchor = self.obj
        args.clear()
        return

    def filter_by_color(self, cell):
        rangos = cell.column[1:, :].visible
        for r in rangos:
            for c in r:
                if c.back_color != cell.back_color:
                    c.rows.visible = False
        return

    def clear(self, what=1023):
        # ~ http://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1CellFlags.html
        self.obj.clearContents(what)
        return

    def transpose(self):
        # ~ 'Flags': 'A',
        # ~ 'FormulaCommand': 0,
        # ~ 'SkipEmptyCells': False,
        # ~ 'AsLink': False,
        # ~ 'MoveMode': 4,
        self.select()
        self.doc.copy()
        self.clear(1023)
        self[0, 0].select()
        self.doc.insert_contents({"Transpose": True})
        _CB.set("")
        return

    def transpose_data(self, formula=False):
        data = self.data
        if formula:
            data = self.formula
        data = tuple(zip(*data))
        self.clear(1023)
        self[0, 0].copy_from(data, formula=formula)
        return

    def merge_by_row(self):
        for r in range(len(self.rows)):
            self[r].merge = True
        return

    def fill(self, source=1):
        self.obj.fillAuto(0, source)
        return

    def _cast(self, t, v):
        if not t:
            return v

        if t == datetime.date:
            nv = datetime.date.fromordinal(int(v) + DATE_OFFSET)
        else:
            nv = t(v)
        return nv

    def get_data(self, types):
        values = [[self._cast(types[i], v) for i, v in enumerate(row)] for row in self.data]
        return values


class LOCalcRanges(object):
    def __init__(self, obj):
        self._obj = obj
        self._ranges = {}
        self._index = 0
        for r in obj:
            sheet = r.Spreadsheet
            rango = LOCalcRange(sheet[r.AbsoluteName])
            self._ranges[rango.name] = rango

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    def __len__(self):
        return self._obj.Count

    def __contains__(self, item):
        return self._obj.hasByName(item.name)

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        try:
            r = self.obj[self._index]
            rango = self._ranges[r.AbsoluteName]
        except IndexError:
            raise StopIteration

        self._index += 1
        return rango

    def __getitem__(self, index):
        if isinstance(index, int):
            r = self.obj[index]
            rango = self._ranges[r.AbsoluteName]
        else:
            rango = self._ranges[index]
        return rango

    @property
    def obj(self):
        return self._obj

    @property
    def names(self):
        return self.obj.ElementNames

    @property
    def data(self):
        return [r.data for r in self._ranges.values()]

    @property
    def style(self):
        return self.obj

    @style.setter
    def style(self, value):
        self.obj.CellStyle = value

    def add(self, rangos):
        if isinstance(rangos, LOCalcRange):
            rangos = (rangos,)
        for r in rangos:
            self._ranges[r.name] = r
            self.obj.addRangeAddress(r.range_address, False)
        return

    def remove(self, rangos):
        if isinstance(rangos, LOCalcRange):
            rangos = (rangos,)
        for r in rangos:
            del self._ranges[r.name]
            self.obj.removeRangeAddress(r.range_address)
        return


class LOWriterStyles(object):
    def __init__(self, styles):
        self._styles = styles

    @property
    def names(self):
        return {s.DisplayName: s.Name for s in self._styles}

    def __str__(self):
        return "\n".join(tuple(self.names.values()))


class LOWriterStylesFamilies(object):
    def __init__(self, styles):
        self._styles = styles

    def __getitem__(self, index):
        styles = {
            "Character": "CharacterStyles",
            "Paragraph": "ParagraphStyles",
            "Page": "PageStyles",
            "Frame": "FrameStyles",
            "Numbering": "NumberingStyles",
            "Table": "TableStyles",
            "Cell": "CellStyles",
        }
        name = styles.get(index, index)
        return LOWriterStyles(self._styles[name])

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        obj = LOWriterStyles(self._styles[self._index])
        self._index += 1
        return obj
        # ~ raise StopIteration

    @property
    def names(self):
        return self._styles.ElementNames

    def __str__(self):
        return "\n".join(self.names)


class LOWriterPageStyle(LOBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    def __str__(self):
        return f"Page Style: {self.name}"

    @property
    def name(self):
        return self._obj.Name


class LOWriterPageStyles(object):
    def __init__(self, styles):
        self._styles = styles

    def __getitem__(self, index):
        return LOWriterPageStyle(self._styles[index])

    @property
    def names(self):
        return self._styles.ElementNames

    def __str__(self):
        return "\n".join(self.names)


class LOWriterTextRange(object):
    def __init__(self, obj, doc):
        self._obj = obj
        self._doc = doc
        self._is_paragraph = self.obj.ImplementationName == "SwXParagraph"
        self._is_table = self.obj.ImplementationName == "SwXTextTable"
        self._is_text = self.obj.ImplementationName == "SwXTextPortion"
        self._is_section = not self.obj.TextSection is None
        self._parts = []
        if self._is_paragraph:
            self._parts = [LOWriterTextRange(p, doc) for p in obj]

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        try:
            obj = self._parts[self._index]
        except IndexError:
            raise StopIteration

        self._index += 1
        return obj

    @property
    def obj(self):
        return self._obj

    @property
    def string(self):
        s = ""
        if not self._is_table:
            s = self.obj.String
        return s

    @string.setter
    def string(self, value):
        self.obj.String = value

    @property
    def value(self):
        return self.string

    @value.setter
    def value(self, value):
        self.string = value

    @property
    def style(self):
        s = ""
        if self.is_paragraph:
            s = self.obj.ParaStyleName
        elif self.is_text:
            s = self.obj.CharStyleName
        return s

    @style.setter
    def style(self, value):
        if self.is_paragraph:
            self.obj.ParaStyleName = value
        elif self.is_text:
            self.obj.CharStyleName = value

    @property
    def is_paragraph(self):
        return self._is_paragraph

    @property
    def is_table(self):
        return self._is_table

    @property
    def is_text(self):
        return self._is_text

    @property
    def is_section(self):
        return self._is_section

    @property
    def text(self):
        return self.obj.Text

    @property
    def cursor(self):
        return self.text.createTextCursorByRange(self.obj)

    @property
    def text_cursor(self):
        return self.text.createTextCursor()

    @property
    def dp(self):
        return self._doc.dp

    @property
    def paragraph(self):
        cursor = self.cursor
        cursor.gotoStartOfParagraph(False)
        cursor.gotoNextParagraph(True)
        return LOWriterTextRange(cursor, self._doc)

    def goto_start(self):
        if self.is_section:
            rango = self.obj.TextSection.Anchor.Start
        else:
            rango = self.obj.Start
        return LOWriterTextRange(rango, self._doc)

    def goto_end(self):
        if self.is_section:
            rango = self.obj.TextSection.Anchor.End
        else:
            rango = self.obj.End
        return LOWriterTextRange(rango, self._doc)

    def goto_previous(self, expand=True):
        cursor = self.cursor
        cursor.gotoPreviousParagraph(expand)
        return LOWriterTextRange(cursor, self._doc)

    def goto_next(self, expand=True):
        cursor = self.cursor
        cursor.gotoNextParagraph(expand)
        return LOWriterTextRange(cursor, self._doc)

    def go_left(self, from_self=True, count=1, expand=False):
        cursor = self.cursor
        if not from_self:
            cursor = self.text_cursor
            cursor.gotoRange(self.obj, False)
        cursor.goLeft(count, expand)
        return LOWriterTextRange(cursor, self._doc)

    def go_right(self, from_self=True, count=1, expand=False):
        cursor = self.cursor
        if not from_self:
            cursor = self.text_cursor
            cursor.gotoRange(self.obj, False)
        cursor.goRight(count, expand)
        return LOWriterTextRange(cursor, self._doc)

    def delete(self):
        self.value = ""
        return

    def offset(self):
        cursor = self.cursor.getEnd()
        return LOWriterTextRange(cursor, self._doc)

    def insert_content(self, data, cursor=None, replace=False):
        if cursor is None:
            cursor = self.cursor
        self.text.insertTextContent(cursor, data, replace)
        return

    def insert_math(self, formula, anchor_type=TextContentAnchorType.AS_CHARACTER, cursor=None, replace=False):
        math = self._doc.create_instance(SERVICES["TEXT_EMBEDDED"])
        math.CLSID = CLSID["FORMULA"]
        math.AnchorType = anchor_type
        self.insert_content(math, cursor, replace)
        math.EmbeddedObject.Component.Formula = formula
        return math

    def new_line(self, count=1):
        cursor = self.cursor
        for i in range(count):
            self.text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
        return LOWriterTextRange(cursor, self._doc)

    def insert_table(self, data):
        table = self._doc.create_instance(SERVICES["TEXT_TABLE"])
        rows = len(data)
        cols = len(data[0])
        table.initialize(rows, cols)
        self.insert_content(table)
        table.DataArray = data
        name = table.Name
        table = LOWriterTextTable(self._doc.tables[name], self._doc)
        return table

    def insert_shape(self, tipo, args={}):
        # ~ args['Width'] = args.get('Width', 1000)
        # ~ args['Height'] = args.get('Height', 1000)
        # ~ args['X'] = args.get('X', 0)
        # ~ args['Y'] = args.get('Y', 0)
        shape = self._doc.dp.add(tipo, args)
        # ~ shape.anchor = self.obj
        return shape

    def insert_image(self, path, args={}):
        w = args.get("Width", 1000)
        h = args.get("Height", 1000)

        image = self._doc.create_instance(SERVICES["GRAPHIC"])
        image.GraphicURL = _P.to_url(path)
        image.AnchorType = TextContentAnchorType.AS_CHARACTER
        image.Width = w
        image.Height = h
        self.insert_content(image)
        return self._doc.dp.last


class LOWriterTextRanges(object):
    def __init__(self, obj, doc):
        self._obj = obj
        self._doc = doc
        self._paragraphs = [LOWriterTextRange(p, doc) for p in obj]

    def __len__(self):
        return len(self._paragraphs)

    def __getitem__(self, index):
        return self._paragraphs[index]

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        try:
            obj = self._paragraphs[self._index]
        except IndexError:
            raise StopIteration

        self._index += 1
        return obj

    @property
    def obj(self):
        return self._obj


class LOWriterTextTable(object):
    def __init__(self, obj, doc):
        self._obj = obj
        self._doc = doc

    @property
    def obj(self):
        return self._obj

    @property
    def name(self):
        return self._obj.Name

    @property
    def data(self):
        return self.obj.DataArray

    @data.setter
    def data(self, values):
        self.obj.DataArray = values

    @property
    def style(self):
        return self.obj.TableTemplateName

    @style.setter
    def style(self, value):
        self.obj.autoFormat(value)


class LOWriterTextTables(object):
    def __init__(self, doc):
        self._doc = doc
        self._obj = doc.obj.TextTables

    def __getitem__(self, key):
        return LOWriterTextTable(self._obj[key], self._doc)

    def __len__(self):
        return self._obj.Count

    def insert(self, data, text_range=None):
        if text_range is None:
            text_range = self._doc.selection
        text_range.insert_table(data)
        return


class LOWriter(LODocument):
    def __init__(self, obj):
        super().__init__(obj)
        self._type = WRITER

    @property
    def text(self):
        return self.paragraphs

    @property
    def paragraphs(self):
        return LOWriterTextRanges(self.obj.Text, self)

    @property
    def tables(self):
        return LOWriterTextTables(self)

    @property
    def selection(self):
        sel = self.obj.CurrentSelection
        if sel.ImplementationName == OBJ_TEXTS:
            if len(sel) == 1:
                sel = LOWriterTextRanges(sel, self)[0]
            else:
                sel = LOWriterTextRanges(sel, self)
            return sel

        if sel.ImplementationName == OBJ_SHAPES:
            if len(sel) == 1:
                sel = sel[0]
                sel = LODrawPage(sel.Parent)[sel.Name]
            return sel

        if sel.ImplementationName == OBJ_GRAPHIC:
            sel = self.dp[sel.Name]
        else:
            debug(sel.ImplementationName)

        return sel

    @property
    def dp(self):
        return self.draw_page

    @property
    def shapes(self):
        return self.draw_page

    @property
    def draw_page(self):
        return LODrawPage(self.obj.DrawPage)

    @property
    def view_cursor(self):
        return self._cc.ViewCursor

    @property
    def cursor(self):
        return self.obj.Text.createTextCursor()

    @property
    def view_cursor(self):
        return self._cc.ViewCursor

    @property
    def page_styles(self):
        ps = self.obj.StyleFamilies["PageStyles"]
        return LOWriterPageStyles(ps)

    @property
    def styles(self):
        return LOWriterStylesFamilies(self.obj.StyleFamilies)

    @property
    def search_descriptor(self):
        return self.obj.createSearchDescriptor()

    @property
    def replace_descriptor(self):
        return self.obj.createReplaceDescriptor()

    @property
    def zoom(self):
        return self._cc.ViewSettings.ZoomValue

    @zoom.setter
    def zoom(self, value):
        self._cc.ViewSettings.ZoomValue = value

    def goto_start(self):
        self.view_cursor.gotoStart(False)
        return self.selection

    def goto_end(self):
        self.view_cursor.gotoEnd(False)
        return self.selection

    def search(self, options, find_all=True):
        descriptor = self.search_descriptor
        descriptor.setSearchString(options.get("Search", ""))
        descriptor.SearchCaseSensitive = options.get("CaseSensitive", False)
        descriptor.SearchWords = options.get("Words", False)
        if "Attributes" in options:
            attr = dict_to_property(options["Attributes"])
            descriptor.setSearchAttributes(attr)
        if hasattr(descriptor, "SearchRegularExpression"):
            descriptor.SearchRegularExpression = options.get("RegularExpression", False)
        if hasattr(descriptor, "SearchType") and "Type" in options:
            descriptor.SearchType = options["Type"]

        result = False
        if find_all:
            found = self.obj.findAll(descriptor)
            if len(found):
                result = [LOWriterTextRange(f, self) for f in found]
        else:
            found = self.obj.findFirst(descriptor)
            if found:
                result = LOWriterTextRange(found, self)

        return result

    def replace(self, options):
        descriptor = self.replace_descriptor
        descriptor.setSearchString(options["Search"])
        descriptor.setReplaceString(options["Replace"])
        descriptor.SearchCaseSensitive = options.get("CaseSensitive", False)
        descriptor.SearchWords = options.get("Words", False)
        if "Attributes" in options:
            attr = dict_to_property(options["Attributes"])
            descriptor.setSearchAttributes(attr)
        if hasattr(descriptor, "SearchRegularExpression"):
            descriptor.SearchRegularExpression = options.get("RegularExpression", False)
        if hasattr(descriptor, "SearchType") and "Type" in options:
            descriptor.SearchType = options["Type"]
        found = self.obj.replaceAll(descriptor)
        return found

    def select(self, text):
        if hasattr(text, "obj"):
            text = text.obj
        self._cc.select(text)
        return


class LOShape(LOBaseObject):
    IMAGE = "com.sun.star.drawing.GraphicObjectShape"

    def __init__(self, obj, index=-1):
        self._index = index
        super().__init__(obj)

    @property
    def type(self):
        t = self.shape_type[21:]
        if self.is_image:
            t = "image"
        return t

    @property
    def shape_type(self):
        return self.obj.ShapeType

    @property
    def properties(self):
        return {}

    @properties.setter
    def properties(self, values):
        _set_properties(self.obj, values)

    @property
    def is_image(self):
        return self.shape_type == self.IMAGE

    @property
    def name(self):
        return self.obj.Name or f"{self.type}{self.index}"

    @name.setter
    def name(self, value):
        self.obj.Name = value

    @property
    def index(self):
        return self._index

    @property
    def size(self):
        s = self.obj.Size
        a = dict(Width=s.Width, Height=s.Height)
        return a

    @property
    def position(self):
        return self.obj.Position

    @property
    def x(self):
        return self.position.X

    @property
    def y(self):
        return self.position.Y

    @property
    def string(self):
        return self.obj.String

    @string.setter
    def string(self, value):
        self.obj.String = value

    @property
    def description(self):
        return self.obj.Description

    @description.setter
    def description(self, value):
        self.obj.Description = value

    @property
    def cell(self):
        return self.anchor

    @property
    def anchor(self):
        obj = self.obj.Anchor
        if obj.ImplementationName == OBJ_CELL:
            obj = LOCalcRange(obj)
        elif obj.ImplementationName == OBJ_TEXT:
            obj = LOWriterTextRange(obj, LODocs().active)
        else:
            debug("Anchor", obj.ImplementationName)
        return obj

    @anchor.setter
    def anchor(self, value):
        if hasattr(value, "obj"):
            value = value.obj
        try:
            self.obj.Anchor = value
        except Exception as e:
            self.obj.AnchorType = value

    @property
    def visible(self):
        return self.obj.Visible

    @visible.setter
    def visible(self, value):
        self.obj.Visible = value

    @property
    def path(self):
        return self.url

    @property
    def url(self):
        url = ""
        if self.is_image:
            url = _P.to_system(self.obj.GraphicURL.OriginURL)
        return url

    @property
    def mimetype(self):
        mt = ""
        if self.is_image:
            mt = self.obj.GraphicURL.MimeType
        return mt

    @property
    def linked(self):
        l = False
        if self.is_image:
            l = self.obj.GraphicURL.Linked
        return l

    def delete(self):
        self.remove()
        return

    def remove(self):
        self.obj.Parent.remove(self.obj)
        return

    def save(self, path: str, mimetype=DEFAULT_MIME_TYPE):
        if _P.is_dir(path):
            name = self.name
            ext = mimetype.lower()
        else:
            p = _P(path)
            path = p.path
            name = p.name
            ext = p.ext.lower()

        path = _P.join(path, f"{name}.{ext}")
        args = dict(
            URL=_P.to_url(path),
            MimeType=MIME_TYPE[ext],
        )
        if not _export_image(self.obj, args):
            path = ""
        return path

    # ~ def save2(self, path: str):
    # ~ size = len(self.obj.Bitmap.DIB)
    # ~ data = self.obj.GraphicStream.readBytes((), size)
    # ~ data = data[-1].value
    # ~ path = _P.join(path, f'{self.name}.png')
    # ~ _P.save_bin(path, b'')
    # ~ return


class LODrawPage(LOBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    def __getitem__(self, index):
        if isinstance(index, int):
            shape = LOShape(self.obj[index], index)
        else:
            for i, o in enumerate(self.obj):
                shape = self.obj[i]
                name = shape.Name or f"shape{i}"
                if name == index:
                    shape = LOShape(shape, i)
                    break
        return shape

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        if self._index == self.count:
            raise StopIteration
        shape = self[self._index]
        self._index += 1
        return shape

    @property
    def name(self):
        return self.obj.Name

    @property
    def doc(self):
        return self.obj.Forms.Parent

    @property
    def width(self):
        return self.obj.Width

    @property
    def height(self):
        return self.obj.Height

    @property
    def count(self):
        return self.obj.Count

    @property
    def last(self):
        return self[self.count - 1]

    def create_instance(self, name):
        return self.doc.createInstance(name)

    def add(self, type_shape, options={}):
        args = options.copy()
        """Insert a shape in page, type shapes:
            Line
            Rectangle
            Ellipse
            Text
            Connector
        """
        index = self.count
        default_height = 3000
        if type_shape == "Line":
            default_height = 0
        w = args.pop("Width", 3000)
        h = args.pop("Height", default_height)
        x = args.pop("X", 1000)
        y = args.pop("Y", 1000)
        name = args.pop("Name", f"{type_shape.lower()}{index}")

        service = f"com.sun.star.drawing.{type_shape}Shape"
        shape = self.create_instance(service)
        shape.Size = Size(w, h)
        shape.Position = Point(x, y)
        shape.Name = name
        self.obj.add(shape)

        if args:
            _set_properties(shape, args)

        return LOShape(self.obj[index], index)

    def remove(self, shape):
        if hasattr(shape, "obj"):
            shape = shape.obj
        return self.obj.remove(shape)

    def remove_all(self):
        while self.count:
            self.obj.remove(self.obj[0])
        return

    def insert_image(self, path, options={}):
        args = options.copy()
        index = self.count
        w = args.get("Width", 3000)
        h = args.get("Height", 3000)
        x = args.get("X", 1000)
        y = args.get("Y", 1000)
        name = args.get("Name", f"image{index}")

        image = self.create_instance("com.sun.star.drawing.GraphicObjectShape")
        if isinstance(path, str):
            image.GraphicURL = _P.to_url(path)
        else:
            gp = create_instance("com.sun.star.graphic.GraphicProvider")
            properties = dict_to_property({"InputStream": path})
            image.Graphic = gp.queryGraphic(properties)

        self.obj.add(image)
        image.Size = Size(w, h)
        image.Position = Point(x, y)
        image.Name = name
        return LOShape(self.obj[index], index)


class LODrawImpress(LODocument):
    def __init__(self, obj):
        super().__init__(obj)

    def __getitem__(self, index):
        if isinstance(index, int):
            page = self.obj.DrawPages[index]
        else:
            page = self.obj.DrawPages.getByName(index)
        return LODrawPage(page)

    @property
    def selection(self):
        sel = self.obj.CurrentSelection[0]
        # ~ return _get_class_uno(sel)
        return sel

    @property
    def current_page(self):
        return LODrawPage(self._cc.getCurrentPage())

    def paste(self):
        call_dispatch(self.frame, ".uno:Paste")
        return self.current_page[-1]

    def add(self, type_shape, args={}):
        return self.current_page.add(type_shape, args)

    def insert_image(self, path, args={}):
        self.current_page.insert_image(path, args)
        return

    # ~ def export(self, path, mimetype='png'):
    # ~ args = dict(
    # ~ URL = _P.to_url(path),
    # ~ MimeType = MIME_TYPE[mimetype],
    # ~ )
    # ~ result = _export_image(self.obj, args)
    # ~ return result


class LODraw(LODrawImpress):
    def __init__(self, obj):
        super().__init__(obj)
        self._type = DRAW


class LOImpress(LODrawImpress):
    def __init__(self, obj):
        super().__init__(obj)
        self._type = IMPRESS


class BaseDateField(DateField):
    def db_value(self, value):
        return _date_to_struct(value)

    def python_value(self, value):
        return _struct_to_date(value)


class BaseTimeField(TimeField):
    def db_value(self, value):
        return _date_to_struct(value)

    def python_value(self, value):
        return _struct_to_date(value)


class BaseDateTimeField(DateTimeField):
    def db_value(self, value):
        return _date_to_struct(value)

    def python_value(self, value):
        return _struct_to_date(value)


class FirebirdDatabase(Database):
    field_types = {"BOOL": "BOOLEAN", "DATETIME": "TIMESTAMP"}

    def __init__(self, database, **kwargs):
        super().__init__(database, **kwargs)
        self._db = database

    def _connect(self):
        return self._db

    def create_tables(self, models, **options):
        options["safe"] = False
        tables = self._db.tables
        models = [m for m in models if not m.__name__.lower() in tables]
        super().create_tables(models, **options)

    def execute_sql(self, sql, params=None, commit=True):
        with __exception_wrapper__:
            cursor = self._db.execute(sql, params)
        return cursor

    def last_insert_id(self, cursor, query_type=None):
        # ~ debug('LAST_ID', cursor)
        return 0

    def rows_affected(self, cursor):
        return self._db.rows_affected

    @property
    def path(self):
        return self._db.path


class BaseRow:
    pass


class BaseQuery(object):
    PY_TYPES = {
        "VARCHAR": "getString",
        "INTEGER": "getLong",
        "DATE": "getDate",
        # ~ 'SQL_LONG': 'getLong',
        # ~ 'SQL_VARYING': 'getString',
        # ~ 'SQL_FLOAT': 'getFloat',
        # ~ 'SQL_BOOLEAN': 'getBoolean',
        # ~ 'SQL_TYPE_DATE': 'getDate',
        # ~ 'SQL_TYPE_TIME': 'getTime',
        # ~ 'SQL_TIMESTAMP': 'getTimestamp',
    }
    # ~ TYPES_DATE = ('SQL_TYPE_DATE', 'SQL_TYPE_TIME', 'SQL_TIMESTAMP')
    TYPES_DATE = ("DATE", "SQL_TYPE_TIME", "SQL_TIMESTAMP")

    def __init__(self, query):
        self._query = query
        self._meta = query.MetaData
        self._cols = self._meta.ColumnCount
        self._names = query.Columns.ElementNames
        self._data = self._get_data()

    def __getitem__(self, index):
        return self._data[index]

    def __iter__(self):
        self._index = 0
        return self

    def __next__(self):
        try:
            row = self._data[self._index]
        except IndexError:
            raise StopIteration
        self._index += 1
        return row

    def _to_python(self, index):
        type_field = self._meta.getColumnTypeName(index)
        # ~ print('TF', type_field)
        value = getattr(self._query, self.PY_TYPES[type_field])(index)
        if type_field in self.TYPES_DATE:
            value = _struct_to_date(value)
        return value

    def _get_row(self):
        row = BaseRow()
        for i in range(1, self._cols + 1):
            column_name = self._meta.getColumnName(i)
            value = self._to_python(i)
            setattr(row, column_name, value)
        return row

    def _get_data(self):
        data = []
        while self._query.next():
            row = self._get_row()
            data.append(row)
        return data

    @property
    def tuples(self):
        data = [tuple(r.__dict__.values()) for r in self._data]
        return tuple(data)

    @property
    def dicts(self):
        data = [r.__dict__ for r in self._data]
        return tuple(data)


class LOBase(object):
    DB_TYPES = {
        str: "setString",
        int: "setInt",
        float: "setFloat",
        bool: "setBoolean",
        Date: "setDate",
        Time: "setTime",
        DateTime: "setTimestamp",
    }
    # ~ setArray
    # ~ setBinaryStream
    # ~ setBlob
    # ~ setByte
    # ~ setBytes
    # ~ setCharacterStream
    # ~ setClob
    # ~ setNull
    # ~ setObject
    # ~ setObjectNull
    # ~ setObjectWithInfo
    # ~ setPropertyValue
    # ~ setRef

    def __init__(self, obj, args={}):
        self._obj = obj
        self._type = BASE
        self._dbc = create_instance("com.sun.star.sdb.DatabaseContext")
        self._rows_affected = 0
        path = args.get("path", "")
        self._path = _P(path)
        self._name = self._path.name
        if _P.exists(path):
            if not self.is_registered:
                self.register()
            db = self._dbc.getByName(self.name)
        else:
            db = self._dbc.createInstance()
            db.URL = "sdbc:embedded:firebird"
            db.DatabaseDocument.storeAsURL(self._path.url, ())
            self.register()
        self._obj = db
        self._con = db.getConnection("", "")

    def __contains__(self, item):
        return item in self.tables

    @property
    def obj(self):
        return self._obj

    @property
    def name(self):
        return self._name

    @property
    def path(self):
        return str(self._path)

    @property
    def is_registered(self):
        return self._dbc.hasRegisteredDatabase(self.name)

    @property
    def tables(self):
        tables = [t.Name.lower() for t in self._con.getTables()]
        return tables

    @property
    def rows_affected(self):
        return self._rows_affected

    def register(self):
        if not self.is_registered:
            self._dbc.registerDatabaseLocation(self.name, self._path.url)
        return

    def revoke(self, name):
        self._dbc.revokeDatabaseLocation(name)
        return True

    def save(self):
        self.obj.DatabaseDocument.store()
        self.refresh()
        return

    def close(self):
        self._con.close()
        return

    def refresh(self):
        self._con.getTables().refresh()
        return

    def initialize(self, database_proxy, tables=[]):
        db = FirebirdDatabase(self)
        database_proxy.initialize(db)
        if tables:
            db.create_tables(tables)
        return

    def _validate_sql(self, sql, params):
        limit = " LIMIT "
        for p in params:
            sql = sql.replace("?", f"'{p}'", 1)
        if limit in sql:
            sql = sql.split(limit)[0]
            sql = sql.replace("SELECT", f"SELECT FIRST {params[-1]}")
        return sql

    def cursor(self, sql, params):
        if sql.startswith("SELECT"):
            sql = self._validate_sql(sql, params)
            cursor = self._con.prepareStatement(sql)
            return cursor

        if not params:
            cursor = self._con.createStatement()
            return cursor

        cursor = self._con.prepareStatement(sql)
        for i, v in enumerate(params, 1):
            t = type(v)
            if not t in self.DB_TYPES:
                error("Type not support")
                debug((i, t, v, self.DB_TYPES[t]))
            getattr(cursor, self.DB_TYPES[t])(i, v)
        return cursor

    def execute(self, sql, params):
        debug(sql, params)
        cursor = self.cursor(sql, params)

        if sql.startswith("SELECT"):
            result = cursor.executeQuery()
        elif params:
            result = cursor.executeUpdate()
            self._rows_affected = result
            self.save()
        else:
            result = cursor.execute(sql)
            self.save()

        return result

    def select(self, sql):
        debug("SELECT", sql)
        if not sql.startswith("SELECT"):
            return ()

        cursor = self._con.prepareStatement(sql)
        query = cursor.executeQuery()
        return BaseQuery(query)

    def get_query(self, query):
        sql, args = query.sql()
        sql = self._validate_sql(sql, args)
        return self.select(sql)


class LOMath(LODocument):
    def __init__(self, obj):
        super().__init__(obj)
        self._type = MATH


class LOBasic(LODocument):
    def __init__(self, obj):
        super().__init__(obj)
        self._type = BASIC


class LODocs(object):
    _desktop = None

    def __init__(self):
        self._desktop = get_desktop()
        LODocs._desktop = self._desktop

    def __getitem__(self, index):
        document = None
        for i, doc in enumerate(self._desktop.Components):
            if isinstance(index, int) and i == index:
                document = _get_class_doc(doc)
                break
            elif isinstance(index, str) and doc.Title == index:
                document = _get_class_doc(doc)
                break
        return document

    def __contains__(self, item):
        doc = self[item]
        return not doc is None

    def __iter__(self):
        self._i = -1
        return self

    def __next__(self):
        self._i += 1
        doc = self[self._i]
        if doc is None:
            raise StopIteration
        else:
            return doc

    def __len__(self):
        # ~ len(self._desktop.Components)
        for i, _ in enumerate(self._desktop.Components):
            pass
        return i + 1

    @property
    def active(self):
        return _get_class_doc(self._desktop.getCurrentComponent())

    @classmethod
    def new(cls, type_doc=CALC, args={}):
        if type_doc == BASE:
            return LOBase(None, args)

        path = f"private:factory/s{type_doc}"
        opt = dict_to_property(args)
        doc = cls._desktop.loadComponentFromURL(path, "_default", 0, opt)
        return _get_class_doc(doc)

    @classmethod
    def open(cls, path, args={}):
        """Open document in path
        Usually options:
            Hidden: True or False
            AsTemplate: True or False
            ReadOnly: True or False
            Password: super_secret
            MacroExecutionMode: 4 = Activate macros
            Preview: True or False

        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html
        http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1MediaDescriptor.html
        """
        path = _P.to_url(path)
        opt = dict_to_property(args)
        doc = cls._desktop.loadComponentFromURL(path, "_default", 0, opt)
        if doc is None:
            return

        return _get_class_doc(doc)

    def connect(self, path):
        db = LOBase(None, {"path": path})
        return db


def _add_listeners(events, control, name=""):
    listeners = {
        "addActionListener": EventsButton,
        "addMouseListener": EventsMouse,
        "addFocusListener": EventsFocus,
        "addItemListener": EventsItem,
        "addKeyListener": EventsKey,
        "addTabListener": EventsTab,
    }
    if hasattr(control, "obj"):
        control = control.obj
    # ~ debug(control.ImplementationName)
    is_grid = control.ImplementationName == "stardiv.Toolkit.GridControl"
    is_link = control.ImplementationName == "stardiv.Toolkit.UnoFixedHyperlinkControl"
    is_roadmap = control.ImplementationName == "stardiv.Toolkit.UnoRoadmapControl"
    is_pages = control.ImplementationName == "stardiv.Toolkit.UnoMultiPageControl"

    for key, value in listeners.items():
        if hasattr(control, key):
            if is_grid and key == "addMouseListener":
                control.addMouseListener(EventsMouseGrid(events, name))
                continue
            if is_link and key == "addMouseListener":
                control.addMouseListener(EventsMouseLink(events, name))
                continue
            if is_roadmap and key == "addItemListener":
                control.addItemListener(EventsItemRoadmap(events, name))
                continue

            getattr(control, key)(listeners[key](events, name))

    if is_grid:
        controllers = EventsGrid(events, name)
        control.addSelectionListener(controllers)
        control.Model.GridDataModel.addGridDataListener(controllers)
    return


def _set_properties(model, properties):
    if "X" in properties:
        properties["PositionX"] = properties.pop("X")
    if "Y" in properties:
        properties["PositionY"] = properties.pop("Y")
    keys = tuple(properties.keys())
    values = tuple(properties.values())
    model.setPropertyValues(keys, values)
    return


class EventsListenerBase(unohelper.Base, XEventListener):
    def __init__(self, controller, name, window=None):
        self._controller = controller
        self._name = name
        self._window = window

    @property
    def name(self):
        return self._name

    def disposing(self, event):
        self._controller = None
        if not self._window is None:
            self._window.setMenuBar(None)


class EventsMouse(EventsListenerBase, XMouseListener, XMouseMotionListener):
    def __init__(self, controller, name):
        super().__init__(controller, name)

    def mousePressed(self, event):
        event_name = "{}_click".format(self._name)
        if event.ClickCount == 2:
            event_name = "{}_double_click".format(self._name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def mouseReleased(self, event):
        pass

    def mouseEntered(self, event):
        pass

    def mouseExited(self, event):
        pass

    # ~ XMouseMotionListener
    def mouseMoved(self, event):
        pass

    def mouseDragged(self, event):
        pass


class EventsMouseLink(EventsMouse):
    def __init__(self, controller, name):
        super().__init__(controller, name)
        self._text_color = 0

    def mouseEntered(self, event):
        model = event.Source.Model
        self._text_color = model.TextColor or 0
        model.TextColor = get_color("blue")
        return

    def mouseExited(self, event):
        model = event.Source.Model
        model.TextColor = self._text_color
        return


class EventsButton(EventsListenerBase, XActionListener):
    def __init__(self, controller, name):
        super().__init__(controller, name)

    def actionPerformed(self, event):
        event_name = f"{self.name}_action"
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return


class EventsFocus(EventsListenerBase, XFocusListener):
    CONTROLS = ("stardiv.Toolkit.UnoControlEditModel",)

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def focusGained(self, event):
        service = event.Source.Model.ImplementationName
        # ~ print('Focus enter', service)
        if service in self.CONTROLS:
            obj = event.Source.Model
            obj.BackgroundColor = COLOR_ON_FOCUS
        return

    def focusLost(self, event):
        service = event.Source.Model.ImplementationName
        if service in self.CONTROLS:
            obj = event.Source.Model
            obj.BackgroundColor = -1
        return


class EventsKey(EventsListenerBase, XKeyListener):
    """
    event.KeyChar
    event.KeyCode
    event.KeyFunc
    event.Modifiers
    """

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def keyPressed(self, event):
        pass

    def keyReleased(self, event):
        event_name = "{}_key_released".format(self._name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        # ~ else:
        # ~ if event.KeyFunc == QUIT and hasattr(self._cls, 'close'):
        # ~ self._cls.close()
        return


class EventsItem(EventsListenerBase, XItemListener):
    def __init__(self, controller, name):
        super().__init__(controller, name)

    def disposing(self, event):
        pass

    def itemStateChanged(self, event):
        event_name = "{}_item_changed".format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return


class EventsItemRoadmap(EventsItem):
    def itemStateChanged(self, event):
        dialog = event.Source.Context.Model
        dialog.Step = event.ItemId + 1
        return


class EventsGrid(EventsListenerBase, XGridDataListener, XGridSelectionListener):
    def __init__(self, controller, name):
        super().__init__(controller, name)

    def dataChanged(self, event):
        event_name = "{}_data_changed".format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def rowHeadingChanged(self, event):
        pass

    def rowsInserted(self, event):
        pass

    def rowsRemoved(self, evemt):
        pass

    def selectionChanged(self, event):
        event_name = "{}_selection_changed".format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return


class EventsMouseGrid(EventsMouse):
    selected = False

    def mousePressed(self, event):
        super().mousePressed(event)
        # ~ obj = event.Source
        # ~ col = obj.getColumnAtPoint(event.X, event.Y)
        # ~ row = obj.getRowAtPoint(event.X, event.Y)
        # ~ print(col, row)
        # ~ if col == -1 and row == -1:
        # ~ if self.selected:
        # ~ obj.deselectAllRows()
        # ~ else:
        # ~ obj.selectAllRows()
        # ~ self.selected = not self.selected
        return

    def mouseReleased(self, event):
        # ~ obj = event.Source
        # ~ col = obj.getColumnAtPoint(event.X, event.Y)
        # ~ row = obj.getRowAtPoint(event.X, event.Y)
        # ~ if row == -1 and col > -1:
        # ~ gdm = obj.Model.GridDataModel
        # ~ for i in range(gdm.RowCount):
        # ~ gdm.updateRowHeading(i, i + 1)
        return


class EventsTab(EventsListenerBase, XTabListener):
    def __init__(self, controller, name):
        super().__init__(controller, name)

    def activated(self, id):
        event_name = "{}_activated".format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(id)
        return


class EventsMenu(EventsListenerBase, XMenuListener):
    def __init__(self, controller):
        super().__init__(controller, "")

    def itemHighlighted(self, event):
        pass

    def itemSelected(self, event):
        name = event.Source.getCommand(event.MenuId)
        if name.startswith("menu"):
            event_name = "{}_selected".format(name)
        else:
            event_name = "menu_{}_selected".format(name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def itemActivated(self, event):
        return

    def itemDeactivated(self, event):
        return


class EventsWindow(EventsListenerBase, XTopWindowListener, XWindowListener):
    def __init__(self, cls):
        self._cls = cls
        super().__init__(cls.events, cls.name, cls._window)

    def windowOpened(self, event):
        event_name = "{}_opened".format(self._name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def windowActivated(self, event):
        control_name = "{}_activated".format(event.Source.Model.Name)
        if hasattr(self._controller, control_name):
            getattr(self._controller, control_name)(event)
        return

    def windowDeactivated(self, event):
        control_name = "{}_deactivated".format(event.Source.Model.Name)
        if hasattr(self._controller, control_name):
            getattr(self._controller, control_name)(event)
        return

    def windowMinimized(self, event):
        pass

    def windowNormalized(self, event):
        pass

    def windowClosing(self, event):
        if self._window:
            control_name = "window_closing"
        else:
            control_name = "{}_closing".format(event.Source.Model.Name)

        if hasattr(self._controller, control_name):
            getattr(self._controller, control_name)(event)
        # ~ else:
        # ~ if not self._modal and not self._block:
        # ~ event.Source.Visible = False
        return

    def windowClosed(self, event):
        control_name = "{}_closed".format(event.Source.Model.Name)
        if hasattr(self._controller, control_name):
            getattr(self._controller, control_name)(event)
        return

    # ~ XWindowListener
    def windowResized(self, event):
        sb = self._cls._subcont
        sb.setPosSize(0, 0, event.Width, event.Height, SIZE)
        event_name = "{}_resized".format(self._name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def windowMoved(self, event):
        pass

    def windowShown(self, event):
        pass

    def windowHidden(self, event):
        pass


# ~ BorderColor = ?
# ~ FontStyleName = ?
# ~ HelpURL = ?
class UnoBaseObject(object):
    def __init__(self, obj, path=""):
        self._obj = obj
        self._model = obj.Model

    def __setattr__(self, name, value):
        exists = hasattr(self, name)
        if not exists and not name in ("_obj", "_model"):
            setattr(self._model, name, value)
        else:
            super().__setattr__(name, value)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def model(self):
        return self._model

    @property
    def m(self):
        return self._model

    @property
    def properties(self):
        return {}

    @properties.setter
    def properties(self, values):
        _set_properties(self.model, values)

    @property
    def name(self):
        return self.model.Name

    @property
    def parent(self):
        return self.obj.Context

    @property
    def tag(self):
        return self.model.Tag

    @tag.setter
    def tag(self, value):
        self.model.Tag = value

    @property
    def visible(self):
        return self.obj.Visible

    @visible.setter
    def visible(self, value):
        self.obj.setVisible(value)

    @property
    def enabled(self):
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value):
        self.model.Enabled = value

    @property
    def step(self):
        return self.model.Step

    @step.setter
    def step(self, value):
        self.model.Step = value

    @property
    def align(self):
        return self.model.Align

    @align.setter
    def align(self, value):
        self.model.Align = value

    @property
    def valign(self):
        return self.model.VerticalAlign

    @valign.setter
    def valign(self, value):
        self.model.VerticalAlign = value

    @property
    def font_weight(self):
        return self.model.FontWeight

    @font_weight.setter
    def font_weight(self, value):
        self.model.FontWeight = value

    @property
    def font_height(self):
        return self.model.FontHeight

    @font_height.setter
    def font_height(self, value):
        self.model.FontHeight = value

    @property
    def font_name(self):
        return self.model.FontName

    @font_name.setter
    def font_name(self, value):
        self.model.FontName = value

    @property
    def font_underline(self):
        return self.model.FontUnderline

    @font_underline.setter
    def font_underline(self, value):
        self.model.FontUnderline = value

    @property
    def text_color(self):
        return self.model.TextColor

    @text_color.setter
    def text_color(self, value):
        self.model.TextColor = value

    @property
    def back_color(self):
        return self.model.BackgroundColor

    @back_color.setter
    def back_color(self, value):
        self.model.BackgroundColor = value

    @property
    def multi_line(self):
        return self.model.MultiLine

    @multi_line.setter
    def multi_line(self, value):
        self.model.MultiLine = value

    @property
    def help_text(self):
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value):
        self.model.HelpText = value

    @property
    def border(self):
        return self.model.Border

    @border.setter
    def border(self, value):
        # ~ Bug for report
        self.model.Border = value

    @property
    def width(self):
        return self._model.Width

    @width.setter
    def width(self, value):
        self.model.Width = value

    @property
    def height(self):
        return self.model.Height

    @height.setter
    def height(self, value):
        self.model.Height = value

    def _get_possize(self, name):
        ps = self.obj.getPosSize()
        return getattr(ps, name)

    def _set_possize(self, name, value):
        ps = self.obj.getPosSize()
        setattr(ps, name, value)
        self.obj.setPosSize(ps.X, ps.Y, ps.Width, ps.Height, POSSIZE)
        return

    @property
    def x(self):
        if hasattr(self.model, "PositionX"):
            return self.model.PositionX
        return self._get_possize("X")

    @x.setter
    def x(self, value):
        if hasattr(self.model, "PositionX"):
            self.model.PositionX = value
        else:
            self._set_possize("X", value)

    @property
    def y(self):
        if hasattr(self.model, "PositionY"):
            return self.model.PositionY
        return self._get_possize("Y")

    @y.setter
    def y(self, value):
        if hasattr(self.model, "PositionY"):
            self.model.PositionY = value
        else:
            self._set_possize("Y", value)

    @property
    def tab_index(self):
        return self._model.TabIndex

    @tab_index.setter
    def tab_index(self, value):
        self.model.TabIndex = value

    @property
    def tab_stop(self):
        return self._model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value):
        self.model.Tabstop = value

    @property
    def ps(self):
        ps = self.obj.getPosSize()
        return ps

    @ps.setter
    def ps(self, ps):
        self.obj.setPosSize(ps.X, ps.Y, ps.Width, ps.Height, POSSIZE)

    def set_focus(self):
        self.obj.setFocus()
        return

    def ps_from(self, source):
        self.ps = source.ps
        return

    def center(self, horizontal=True, vertical=False):
        p = self.parent.Model
        w = p.Width
        h = p.Height
        if horizontal:
            x = w / 2 - self.width / 2
            self.x = x
        if vertical:
            y = h / 2 - self.height / 2
            self.y = y
        return

    def move(self, origin, x=0, y=5, center=False):
        if x:
            self.x = origin.x + origin.width + x
        else:
            self.x = origin.x
        if y:
            self.y = origin.y + origin.height + y
        else:
            self.y = origin.y

        if center:
            self.center()
        return


class UnoLabel(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "label"

    @property
    def value(self):
        return self.model.Label

    @value.setter
    def value(self, value):
        self.model.Label = value


class UnoLabelLink(UnoLabel):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "link"


class UnoButton(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "button"

    @property
    def value(self):
        return self.model.Label

    @value.setter
    def value(self, value):
        self.model.Label = value


class UnoRadio(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "radio"

    @property
    def value(self):
        return self.model.Label

    @value.setter
    def value(self, value):
        self.model.Label = value


class UnoCheckBox(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "checkbox"

    @property
    def value(self):
        return self.model.State

    @value.setter
    def value(self, value):
        self.model.State = value

    @property
    def label(self):
        return self.model.Label

    @label.setter
    def label(self, value):
        self.model.Label = value

    @property
    def tri_state(self):
        return self.model.TriState

    @tri_state.setter
    def tri_state(self, value):
        self.model.TriState = value


# ~ https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlEditModel.html
class UnoText(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "text"

    @property
    def value(self):
        return self.model.Text

    @value.setter
    def value(self, value):
        self.model.Text = value

    @property
    def echochar(self):
        return chr(self.model.EchoChar)

    @echochar.setter
    def echochar(self, value):
        self.model.EchoChar = ord(value[0])

    def validate(self):
        return


class UnoImage(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return "image"

    @property
    def value(self):
        return self.url

    @value.setter
    def value(self, value):
        self.url = value

    @property
    def url(self):
        return self.m.ImageURL

    @url.setter
    def url(self, value):
        self.m.ImageURL = None
        self.m.ImageURL = _P.to_url(value)


class UnoListBox(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)
        self._path = ""

    def __setattr__(self, name, value):
        if name in ("_path",):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    @property
    def type(self):
        return "listbox"

    @property
    def value(self):
        return self.obj.getSelectedItem()

    @property
    def count(self):
        return len(self.data)

    @property
    def data(self):
        return self.model.StringItemList

    @data.setter
    def data(self, values):
        self.model.StringItemList = list(sorted(values))

    @property
    def path(self):
        return self._path

    @path.setter
    def path(self, value):
        self._path = value

    def unselect(self):
        self.obj.selectItem(self.value, False)
        return

    def select(self, pos=0):
        if isinstance(pos, str):
            self.obj.selectItem(pos, True)
        else:
            self.obj.selectItemPos(pos, True)
        return

    def clear(self):
        self.model.removeAllItems()
        return

    def _set_image_url(self, image):
        if _P.exists(image):
            return _P.to_url(image)

        path = _P.join(self._path, DIR["images"], image)
        return _P.to_url(path)

    def insert(self, value, path="", pos=-1, show=True):
        if pos < 0:
            pos = self.count
        if path:
            self.model.insertItem(pos, value, self._set_image_url(path))
        else:
            self.model.insertItemText(pos, value)
        if show:
            self.select(pos)
        return


class UnoRoadmap(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)
        self._options = ()

    def __setattr__(self, name, value):
        if name in ("_options",):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    @property
    def options(self):
        return self._options

    @options.setter
    def options(self, values):
        self._options = values
        for i, v in enumerate(values):
            opt = self.model.createInstance()
            opt.ID = i
            opt.Label = v
            self.model.insertByIndex(i, opt)
        return

    @property
    def enabled(self):
        return True

    @enabled.setter
    def enabled(self, value):
        for m in self.model:
            m.Enabled = value
        return

    def set_enabled(self, index, value):
        self.model.getByIndex(index).Enabled = value
        return


class UnoTree(UnoBaseObject):
    def __init__(
        self,
        obj,
    ):
        super().__init__(obj)
        self._tdm = None
        self._data = []

    def __setattr__(self, name, value):
        if name in ("_tdm", "_data"):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    @property
    def selection(self):
        sel = self.obj.Selection
        return sel.DataValue, sel.DisplayValue

    @property
    def parent(self):
        parent = self.obj.Selection.Parent
        if parent is None:
            return ()
        return parent.DataValue, parent.DisplayValue

    def _get_parents(self, node):
        value = (node.DisplayValue,)
        parent = node.Parent
        if parent is None:
            return value
        return self._get_parents(parent) + value

    @property
    def parents(self):
        values = self._get_parents(self.obj.Selection)
        return values

    @property
    def root(self):
        if self._tdm is None:
            return ""
        return self._tdm.Root.DisplayValue

    @root.setter
    def root(self, value):
        self._add_data_model(value)

    def _add_data_model(self, name):
        tdm = create_instance("com.sun.star.awt.tree.MutableTreeDataModel")
        root = tdm.createNode(name, True)
        root.DataValue = 0
        tdm.setRoot(root)
        self.model.DataModel = tdm
        self._tdm = self.model.DataModel
        return

    @property
    def path(self):
        return self.root

    @path.setter
    def path(self, value):
        self.data = _P.walk_dir(value, True)

    @property
    def data(self):
        return self._data

    @data.setter
    def data(self, values):
        self._data = list(values)
        self._add_data()

    def _add_data(self):
        if not self.data:
            return

        parents = {}
        for node in self.data:
            parent = parents.get(node[1], self._tdm.Root)
            child = self._tdm.createNode(node[2], False)
            child.DataValue = node[0]
            parent.appendChild(child)
            parents[node[0]] = child
        self.obj.expandNode(self._tdm.Root)
        return


# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1grid.html
class UnoGrid(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)
        self._gdm = self.model.GridDataModel
        self._data = []
        self._formats = ()

    def __setattr__(self, name, value):
        if name in ("_gdm", "_data", "_formats"):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    def __getitem__(self, key):
        value = self._gdm.getCellData(key[0], key[1])
        return value

    def __setitem__(self, key, value):
        self._gdm.updateCellData(key[0], key[1], value)
        return

    @property
    def type(self):
        return "grid"

    @property
    def columns(self):
        return {}

    @columns.setter
    def columns(self, values):
        # ~ self._columns = values
        # ~ https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1grid_1_1XGridColumn.html
        model = create_instance("com.sun.star.awt.grid.DefaultGridColumnModel", True)
        for properties in values:
            column = create_instance("com.sun.star.awt.grid.GridColumn", True)
            for k, v in properties.items():
                setattr(column, k, v)
            model.addColumn(column)
        self.model.ColumnModel = model
        return

    @property
    def data(self):
        return self._data

    @data.setter
    def data(self, values):
        self._data = values
        self.clear()
        headings = tuple(range(1, len(values) + 1))
        self._gdm.addRows(headings, values)
        # ~ rows = range(grid_dm.RowCount)
        # ~ colors = [COLORS['GRAY'] if r % 2 else COLORS['WHITE'] for r in rows]
        # ~ grid.Model.RowBackgroundColors = tuple(colors)
        return

    @property
    def value(self):
        if self.column == -1 or self.row == -1:
            return ""
        return self[self.column, self.row]

    @value.setter
    def value(self, value):
        if self.column > -1 and self.row > -1:
            self[self.column, self.row] = value

    @property
    def row(self):
        return self.obj.CurrentRow

    @property
    def row_count(self):
        return self._gdm.RowCount

    @property
    def column(self):
        return self.obj.CurrentColumn

    @property
    def column(self):
        return self.obj.CurrentColumn

    @property
    def is_valid(self):
        return not (self.row == -1 or self.column == -1)

    @property
    def formats(self):
        return self._formats

    @formats.setter
    def formats(self, values):
        self._formats = values

    def clear(self):
        self._gdm.removeAllRows()
        return

    def _format_columns(self, data):
        row = data
        if self.formats:
            for i, f in enumerate(formats):
                if f:
                    row[i] = f.format(data[i])
        return row

    def add_row(self, data):
        self._data.append(data)
        row = self._format_columns(data)
        self._gdm.addRow(self.row_count + 1, row)
        return

    def set_cell_tooltip(self, col, row, value):
        self._gdm.updateCellToolTip(col, row, value)
        return

    def get_cell_tooltip(self, col, row):
        value = self._gdm.getCellToolTip(col, row)
        return value

    def sort(self, column, asc=True):
        self._gdm.sortByColumn(column, asc)
        self.update_row_heading()
        return

    def update_row_heading(self):
        for i in range(self.row_count):
            self._gdm.updateRowHeading(i, i + 1)
        return

    def remove_row(self, row):
        self._gdm.removeRow(row)
        del self._data[row]
        self.update_row_heading()
        return


class UnoPage(object):
    def __init__(self, obj):
        self._obj = obj
        self._events = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def model(self):
        return self._obj.Model

    # ~ @property
    # ~ def id(self):
    # ~ return self.m.TabPageID

    @property
    def parent(self):
        return self.obj.Context

    def _set_image_url(self, image):
        if _P.exists(image):
            return _P.to_url(image)

        path = _P.join(self._path, DIR["images"], image)
        return _P.to_url(path)

    def _special_properties(self, tipo, args):
        if tipo == "link" and not "Label" in args:
            args["Label"] = args["URL"]
            return args

        if tipo == "button":
            if "ImageURL" in args:
                args["ImageURL"] = self._set_image_url(args["ImageURL"])
            args["FocusOnClick"] = args.get("FocusOnClick", False)
            return args

        if tipo == "roadmap":
            args["Height"] = args.get("Height", self.height)
            if "Title" in args:
                args["Text"] = args.pop("Title")
            return args

        if tipo == "tree":
            args["SelectionType"] = args.get("SelectionType", SINGLE)
            return args

        if tipo == "grid":
            args["ShowRowHeader"] = args.get("ShowRowHeader", True)
            return args

        if tipo == "pages":
            args["Width"] = args.get("Width", self.width)
            args["Height"] = args.get("Height", self.height)

        return args

    def add_control(self, args):
        tipo = args.pop("Type").lower()
        root = args.pop("Root", "")
        sheets = args.pop("Sheets", ())
        columns = args.pop("Columns", ())

        args = self._special_properties(tipo, args)
        model = self.model.createInstance(UNO_MODELS[tipo])
        _set_properties(model, args)
        name = args["Name"]
        self.model.insertByName(name, model)
        control = self.obj.getControl(name)
        _add_listeners(self._events, control, name)
        control = UNO_CLASSES[tipo](control)

        if tipo in ("listbox",):
            control.path = self.path

        if tipo == "tree" and root:
            control.root = root
        elif tipo == "grid" and columns:
            control.columns = columns
        elif tipo == "pages" and sheets:
            control.sheets = sheets
            control.events = self.events

        setattr(self, name, control)
        return control


class UnoPages(UnoBaseObject):
    def __init__(self, obj):
        super().__init__(obj)
        self._sheets = []
        self._events = None

    def __setattr__(self, name, value):
        if name in ("_sheets", "_events"):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    def __getitem__(self, index):
        name = index
        if isinstance(index, int):
            name = f"sheet{index}"
        sheet = self.obj.getControl(name)
        page = UnoPage(sheet)
        page._events = self._events
        return page

    @property
    def type(self):
        return "pages"

    @property
    def current(self):
        return self.obj.ActiveTabID

    @property
    def active(self):
        return self.current

    @property
    def sheets(self):
        return self._sheets

    @sheets.setter
    def sheets(self, values):
        self._sheets = values
        for i, title in enumerate(values):
            sheet = self.m.createInstance("com.sun.star.awt.UnoPageModel")
            sheet.Title = title
            self.m.insertByName(f"sheet{i + 1}", sheet)
        return

    @property
    def events(self):
        return self._events

    @events.setter
    def events(self, controllers):
        self._events = controllers

    @property
    def visible(self):
        return self.obj.Visible

    @visible.setter
    def visible(self, value):
        self.obj.Visible = value

    def insert(self, title):
        self._sheets.append(title)
        id = len(self._sheets)
        sheet = self.m.createInstance("com.sun.star.awt.UnoPageModel")
        sheet.Title = title
        self.m.insertByName(f"sheet{id}", sheet)
        return self[id]

    def remove(self, id):
        self.obj.removeTab(id)
        return

    def activate(self, id):
        self.obj.activateTab(id)
        return


UNO_CLASSES = {
    "label": UnoLabel,
    "link": UnoLabelLink,
    "button": UnoButton,
    "radio": UnoRadio,
    "checkbox": UnoCheckBox,
    "text": UnoText,
    "image": UnoImage,
    "listbox": UnoListBox,
    "roadmap": UnoRoadmap,
    "tree": UnoTree,
    "grid": UnoGrid,
    "pages": UnoPages,
}

UNO_MODELS = {
    "label": "com.sun.star.awt.UnoControlFixedTextModel",
    "link": "com.sun.star.awt.UnoControlFixedHyperlinkModel",
    "button": "com.sun.star.awt.UnoControlButtonModel",
    "radio": "com.sun.star.awt.UnoControlRadioButtonModel",
    "checkbox": "com.sun.star.awt.UnoControlCheckBoxModel",
    "text": "com.sun.star.awt.UnoControlEditModel",
    "image": "com.sun.star.awt.UnoControlImageControlModel",
    "listbox": "com.sun.star.awt.UnoControlListBoxModel",
    "roadmap": "com.sun.star.awt.UnoControlRoadmapModel",
    "tree": "com.sun.star.awt.tree.TreeControlModel",
    "grid": "com.sun.star.awt.grid.UnoControlGridModel",
    "pages": "com.sun.star.awt.UnoMultiPageModel",
    "groupbox": "com.sun.star.awt.UnoControlGroupBoxModel",
    "combobox": "com.sun.star.awt.UnoControlComboBoxModel",
}
# ~ 'CurrencyField': 'com.sun.star.awt.UnoControlCurrencyFieldModel',
# ~ 'DateField': 'com.sun.star.awt.UnoControlDateFieldModel',
# ~ 'FileControl': 'com.sun.star.awt.UnoControlFileControlModel',
# ~ 'FormattedField': 'com.sun.star.awt.UnoControlFormattedFieldModel',
# ~ 'NumericField': 'com.sun.star.awt.UnoControlNumericFieldModel',
# ~ 'PatternField': 'com.sun.star.awt.UnoControlPatternFieldModel',
# ~ 'ProgressBar': 'com.sun.star.awt.UnoControlProgressBarModel',
# ~ 'ScrollBar': 'com.sun.star.awt.UnoControlScrollBarModel',
# ~ 'SimpleAnimation': 'com.sun.star.awt.UnoControlSimpleAnimationModel',
# ~ 'SpinButton': 'com.sun.star.awt.UnoControlSpinButtonModel',
# ~ 'Throbber': 'com.sun.star.awt.UnoControlThrobberModel',
# ~ 'TimeField': 'com.sun.star.awt.UnoControlTimeFieldModel',


class LODialog(object):
    SEPARATION = 5
    MODELS = {
        "label": "com.sun.star.awt.UnoControlFixedTextModel",
        "link": "com.sun.star.awt.UnoControlFixedHyperlinkModel",
        "button": "com.sun.star.awt.UnoControlButtonModel",
        "radio": "com.sun.star.awt.UnoControlRadioButtonModel",
        "checkbox": "com.sun.star.awt.UnoControlCheckBoxModel",
        "text": "com.sun.star.awt.UnoControlEditModel",
        "image": "com.sun.star.awt.UnoControlImageControlModel",
        "listbox": "com.sun.star.awt.UnoControlListBoxModel",
        "roadmap": "com.sun.star.awt.UnoControlRoadmapModel",
        "tree": "com.sun.star.awt.tree.TreeControlModel",
        "grid": "com.sun.star.awt.grid.UnoControlGridModel",
        "pages": "com.sun.star.awt.UnoMultiPageModel",
        "groupbox": "com.sun.star.awt.UnoControlGroupBoxModel",
        "combobox": "com.sun.star.awt.UnoControlComboBoxModel",
    }

    def __init__(self, args):
        self._obj = self._create(args)
        self._model = self.obj.Model
        self._events = None
        self._modal = True
        self._controls = {}
        self._color_on_focus = COLOR_ON_FOCUS
        self._id = ""
        self._path = ""
        self._init_controls()

    def _create(self, args):
        service = "com.sun.star.awt.DialogProvider"
        path = args.pop("Path", "")
        if path:
            dp = create_instance(service, True)
            dlg = dp.createDialog(_P.to_url(path))
            return dlg

        if "Location" in args:
            name = args["Name"]
            library = args.get("Library", "Standard")
            location = args.get("Location", "application").lower()
            if location == "user":
                location = "application"
            url = f"vnd.sun.star.script:{library}.{name}?location={location}"
            if location == "document":
                dp = create_instance(service, args=docs.active.obj)
            else:
                dp = create_instance(service, True)
                # ~ uid = docs.active.uid
                # ~ url = f'vnd.sun.star.tdoc:/{uid}/Dialogs/{library}/{name}.xml'
            dlg = dp.createDialog(url)
            return dlg

        dlg = create_instance("com.sun.star.awt.UnoControlDialog", True)
        model = create_instance("com.sun.star.awt.UnoControlDialogModel", True)
        toolkit = create_instance("com.sun.star.awt.Toolkit", True)
        _set_properties(model, args)
        dlg.setModel(model)
        dlg.setVisible(False)
        dlg.createPeer(toolkit, None)
        return dlg

    def _get_type_control(self, name):
        name = name.split(".")[2]
        types = {
            "UnoFixedTextControl": "label",
            "UnoEditControl": "text",
            "UnoButtonControl": "button",
        }
        return types[name]

    def _init_controls(self):
        for control in self.obj.getControls():
            tipo = self._get_type_control(control.ImplementationName)
            name = control.Model.Name
            control = UNO_CLASSES[tipo](control)
            setattr(self, name, control)
        return

    @property
    def obj(self):
        return self._obj

    @property
    def model(self):
        return self._model

    @property
    def controls(self):
        return self._controls

    @property
    def path(self):
        return self._path

    @property
    def id(self):
        return self._id

    @id.setter
    def id(self, value):
        self._id = value
        self._path = _P.from_id(value)

    @property
    def height(self):
        return self.model.Height

    @height.setter
    def height(self, value):
        self.model.Height = value

    @property
    def width(self):
        return self.model.Width

    @width.setter
    def width(self, value):
        self.model.Width = value

    @property
    def visible(self):
        return self.obj.Visible

    @visible.setter
    def visible(self, value):
        self.obj.Visible = value

    @property
    def step(self):
        return self.model.Step

    @step.setter
    def step(self, value):
        self.model.Step = value

    @property
    def events(self):
        return self._events

    @events.setter
    def events(self, controllers):
        self._events = controllers(self)
        self._connect_listeners()

    @property
    def color_on_focus(self):
        return self._color_on_focus

    @color_on_focus.setter
    def color_on_focus(self, value):
        self._color_on_focus = get_color(value)

    def _connect_listeners(self):
        for control in self.obj.Controls:
            _add_listeners(self.events, control, control.Model.Name)
        return

    def _set_image_url(self, image):
        if _P.exists(image):
            return _P.to_url(image)

        path = _P.join(self._path, DIR["images"], image)
        return _P.to_url(path)

    def _special_properties(self, tipo, args):
        if tipo == "link" and not "Label" in args:
            args["Label"] = args["URL"]
            return args

        if tipo == "button":
            if "ImageURL" in args:
                args["ImageURL"] = self._set_image_url(args["ImageURL"])
            args["FocusOnClick"] = args.get("FocusOnClick", False)
            return args

        if tipo == "roadmap":
            args["Height"] = args.get("Height", self.height)
            if "Title" in args:
                args["Text"] = args.pop("Title")
            return args

        if tipo == "tree":
            args["SelectionType"] = args.get("SelectionType", SINGLE)
            return args

        if tipo == "grid":
            args["ShowRowHeader"] = args.get("ShowRowHeader", True)
            return args

        if tipo == "pages":
            args["Width"] = args.get("Width", self.width)
            args["Height"] = args.get("Height", self.height)

        return args

    def add_control(self, args):
        tipo = args.pop("Type").lower()
        root = args.pop("Root", "")
        sheets = args.pop("Sheets", ())
        columns = args.pop("Columns", ())

        args = self._special_properties(tipo, args)
        model = self.model.createInstance(self.MODELS[tipo])
        _set_properties(model, args)
        name = args["Name"]
        self.model.insertByName(name, model)
        control = self.obj.getControl(name)
        _add_listeners(self.events, control, name)
        control = UNO_CLASSES[tipo](control)

        if tipo in ("listbox",):
            control.path = self.path

        if tipo == "tree" and root:
            control.root = root
        elif tipo == "grid" and columns:
            control.columns = columns
        elif tipo == "pages" and sheets:
            control.sheets = sheets
            control.events = self.events

        setattr(self, name, control)
        self._controls[name] = control
        return control

    def center(self, control, x=0, y=0):
        w = self.width
        h = self.height

        if isinstance(control, tuple):
            wt = self.SEPARATION * -1
            for c in control:
                wt += c.width + self.SEPARATION
            x = w / 2 - wt / 2
            for c in control:
                c.x = x
                x = c.x + c.width + self.SEPARATION
            return

        if x < 0:
            x = w + x - control.width
        elif x == 0:
            x = w / 2 - control.width / 2
        if y < 0:
            y = h + y - control.height
        elif y == 0:
            y = h / 2 - control.height / 2
        control.x = x
        control.y = y
        return

    def open(self, modal=True):
        self._modal = modal
        if modal:
            return self.obj.execute()
        else:
            self.visible = True
        return

    def close(self, value=0):
        if self._modal:
            value = self.obj.endDialog(value)
        else:
            self.visible = False
            self.obj.dispose()
        return value

    def set_values(self, data):
        for k, v in data.items():
            self._controls[k].value = v
        return


class LOSheets(object):
    def __getitem__(self, index):
        return LODocs().active[index]

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass


class LOCells(object):
    def __getitem__(self, index):
        return LODocs().active.active[index]


class LOShortCut(object):
    # ~ getKeyEventsByCommand

    def __init__(self, app):
        self._app = app
        self._scm = None
        self._init_values()

    def _init_values(self):
        name = "com.sun.star.ui.GlobalAcceleratorConfiguration"
        instance = "com.sun.star.ui.ModuleUIConfigurationManagerSupplier"
        service = TYPE_DOC[self._app]
        manager = create_instance(instance, True)
        uicm = manager.getUIConfigurationManager(service)
        self._scm = uicm.ShortCutManager
        return

    def __contains__(self, item):
        cmd = self._get_command(item)
        return bool(cmd)

    def _get_key_event(self, command):
        events = self._scm.AllKeyEvents
        for event in events:
            cmd = self._scm.getCommandByKeyEvent(event)
            if cmd == command:
                break
        return event

    def _to_key_event(self, shortcut):
        key_event = KeyEvent()
        keys = shortcut.split("+")
        for v in keys[:-1]:
            key_event.Modifiers += MODIFIERS[v.lower()]
        key_event.KeyCode = getattr(Key, keys[-1].upper())
        return key_event

    def _get_command(self, shortcut):
        command = ""
        key_event = self._to_key_event(shortcut)
        try:
            command = self._scm.getCommandByKeyEvent(key_event)
        except NoSuchElementException:
            debug(f"No exists: {shortcut}")
        return command

    def add(self, shortcut, command):
        if isinstance(command, dict):
            command = _get_url_script(command)
        key_event = self._to_key_event(shortcut)
        self._scm.setKeyEvent(key_event, command)
        self._scm.store()
        return

    def reset(self):
        self._scm.reset()
        self._scm.store()
        return

    def remove(self, shortcut):
        key_event = self._to_key_event(shortcut)
        try:
            self._scm.removeKeyEvent(key_event)
            self._scm.store()
        except NoSuchElementException:
            debug(f"No exists: {shortcut}")
        return

    def remove_by_command(self, command):
        if isinstance(command, dict):
            command = _get_url_script(command)
        try:
            self._scm.removeCommandFromAllKeyEvents(command)
            self._scm.store()
        except NoSuchElementException:
            debug(f"No exists: {command}")
        return


class LOShortCuts(object):
    def __getitem__(self, index):
        return LOShortCut(index)


class LOMenu(object):
    def __init__(self, app):
        self._app = app
        self._ui = None
        self._pymenus = None
        self._menu = None
        self._menus = self._get_menus()

    def __getitem__(self, index):
        if isinstance(index, int):
            self._menu = self._menus[index]
        else:
            for menu in self._menus:
                cmd = menu.get("CommandURL", "")
                if MENUS[index.lower()] == cmd:
                    self._menu = menu
                    break
        # ~ line = self._menu.get('CommandURL', '')
        # ~ line += self._get_submenus(self._menu['ItemDescriptorContainer'])
        return self._menu

    def _get_menus(self):
        instance = "com.sun.star.ui.ModuleUIConfigurationManagerSupplier"
        service = TYPE_DOC[self._app]
        manager = create_instance(instance, True)
        self._ui = manager.getUIConfigurationManager(service)
        self._pymenus = self._ui.getSettings(NODE_MENUBAR, True)
        data = []
        for menu in self._pymenus:
            data.append(data_to_dict(menu))
        return data

    def _get_info(self, menu):
        line = menu.get("CommandURL", "")
        line += self._get_submenus(menu["ItemDescriptorContainer"])
        return line

    def _get_submenus(self, menu, level=1):
        line = ""
        for i, v in enumerate(menu):
            data = data_to_dict(v)
            cmd = data.get("CommandURL", "----------")
            line += f'\n{"  " * level}├─ ({i}) {cmd}'
            submenu = data.get("ItemDescriptorContainer", None)
            if not submenu is None:
                line += self._get_submenus(submenu, level + 1)
        return line

    def __str__(self):
        info = "\n".join([self._get_info(m) for m in self._menus])
        return info

    def _get_index_menu(self, menu, command):
        index = -1
        for i, v in enumerate(menu):
            data = data_to_dict(v)
            cmd = data.get("CommandURL", "")
            if cmd == command:
                index = i
                break
        return index

    def insert(self, name, args):
        idc = None
        replace = False
        command = args["CommandURL"]
        label = args["Label"]

        self[name]
        menu = self._menu["ItemDescriptorContainer"]
        submenu = args.get("Submenu", False)
        if submenu:
            idc = self._ui.createSettings()

        index = self._get_index_menu(menu, command)
        if index == -1:
            if "Index" in args:
                index = args["Index"]
            else:
                index = self._get_index_menu(menu, args["After"]) + 1
        else:
            replace = True

        data = dict(
            CommandURL=command,
            Label=label,
            Style=0,
            Type=0,
            ItemDescriptorContainer=idc,
        )
        self._save(menu, data, index, replace)
        self._insert_submenu(idc, submenu)
        return

    def _get_command(self, args):
        shortcut = args.get("ShortCut", "")
        cmd = args["CommandURL"]
        if isinstance(cmd, dict):
            cmd = _get_url_script(cmd)
        if shortcut:
            LOShortCut(self._app).add(shortcut, cmd)
        return cmd

    def _insert_submenu(self, parent, menus):
        for i, v in enumerate(menus):
            submenu = v.pop("Submenu", False)
            if submenu:
                idc = self._ui.createSettings()
                v["ItemDescriptorContainer"] = idc
            v["Type"] = 0
            if v["Label"] == "-":
                v["Type"] = 1
            else:
                v["CommandURL"] = self._get_command(v)
            self._save(parent, v, i)
            if submenu:
                self._insert_submenu(idc, submenu)
        return

    def remove(self, name, command):
        self[name]
        menu = self._menu["ItemDescriptorContainer"]
        index = self._get_index_menu(menu, command)
        if index > -1:
            uno.invoke(menu, "removeByIndex", (index,))
            self._ui.replaceSettings(NODE_MENUBAR, self._pymenus)
            self._ui.store()
        return

    def _save(self, menu, properties, index, replace=False):
        properties = dict_to_property(properties, True)
        if replace:
            uno.invoke(menu, "replaceByIndex", (index, properties))
        else:
            uno.invoke(menu, "insertByIndex", (index, properties))
        self._ui.replaceSettings(NODE_MENUBAR, self._pymenus)
        self._ui.store()
        return


class LOMenus(object):
    def __getitem__(self, index):
        return LOMenu(index)


class LOWindow(object):
    EMPTY = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE dlg:window PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "dialog.dtd">
<dlg:window xmlns:dlg="http://openoffice.org/2000/dialog" xmlns:script="http://openoffice.org/2000/script" dlg:id="empty" dlg:left="0" dlg:top="0" dlg:width="0" dlg:height="0" dlg:closeable="true" dlg:moveable="true" dlg:withtitlebar="false"/>"""
    MODELS = {
        "label": "com.sun.star.awt.UnoControlFixedTextModel",
        "link": "com.sun.star.awt.UnoControlFixedHyperlinkModel",
        "button": "com.sun.star.awt.UnoControlButtonModel",
        "radio": "com.sun.star.awt.UnoControlRadioButtonModel",
        "checkbox": "com.sun.star.awt.UnoControlCheckBoxModel",
        "text": "com.sun.star.awt.UnoControlEditModel",
        "image": "com.sun.star.awt.UnoControlImageControlModel",
        "listbox": "com.sun.star.awt.UnoControlListBoxModel",
        "roadmap": "com.sun.star.awt.UnoControlRoadmapModel",
        "tree": "com.sun.star.awt.tree.TreeControlModel",
        "grid": "com.sun.star.awt.grid.UnoControlGridModel",
        "pages": "com.sun.star.awt.UnoMultiPageModel",
        "groupbox": "com.sun.star.awt.UnoControlGroupBoxModel",
        "combobox": "com.sun.star.awt.UnoControlComboBoxModel",
    }

    def __init__(self, args):
        self._events = None
        self._menu = None
        self._container = None
        self._model = None
        self._id = ""
        self._path = ""
        self._obj = self._create(args)

    def _create(self, properties):
        ps = (
            properties.get("X", 0),
            properties.get("Y", 0),
            properties.get("Width", 500),
            properties.get("Height", 500),
        )
        self._title = properties.get("Title", TITLE)
        self._create_frame(ps)
        self._create_container(ps)
        self._create_subcontainer(ps)
        # ~ self._create_splitter(ps)
        return

    def _create_frame(self, ps):
        service = "com.sun.star.frame.TaskCreator"
        tc = create_instance(service, True)
        self._frame = tc.createInstanceWithArguments(
            (
                NamedValue("FrameName", "EasyMacroWin"),
                NamedValue("PosSize", Rectangle(*ps)),
            )
        )
        self._window = self._frame.getContainerWindow()
        self._toolkit = self._window.getToolkit()
        desktop = get_desktop()
        self._frame.setCreator(desktop)
        desktop.getFrames().append(self._frame)
        self._frame.Title = self._title
        return

    def _create_container(self, ps):
        service = "com.sun.star.awt.UnoControlContainer"
        self._container = create_instance(service, True)
        service = "com.sun.star.awt.UnoControlContainerModel"
        model = create_instance(service, True)
        model.BackgroundColor = get_color((225, 225, 225))
        self._container.setModel(model)
        self._container.createPeer(self._toolkit, self._window)
        self._container.setPosSize(*ps, POSSIZE)
        self._frame.setComponent(self._container, None)
        return

    def _create_subcontainer(self, ps):
        service = "com.sun.star.awt.ContainerWindowProvider"
        cwp = create_instance(service, True)

        path_tmp = _P.save_tmp(self.EMPTY)
        subcont = cwp.createContainerWindow(_P.to_url(path_tmp), "", self._container.getPeer(), None)
        _P.kill(path_tmp)

        subcont.setPosSize(0, 0, 500, 500, POSSIZE)
        subcont.setVisible(True)
        self._container.addControl("subcont", subcont)
        self._subcont = subcont
        self._model = subcont.Model
        return

    def _create_popupmenu(self, menus):
        menu = create_instance("com.sun.star.awt.PopupMenu", True)
        for i, m in enumerate(menus):
            label = m["label"]
            cmd = m.get("event", "")
            if not cmd:
                cmd = label.lower().replace(" ", "_")
            if label == "-":
                menu.insertSeparator(i)
            else:
                menu.insertItem(i, label, m.get("style", 0), i)
                menu.setCommand(i, cmd)
                # ~ menu.setItemImage(i, path?, True)
        menu.addMenuListener(EventsMenu(self.events))
        return menu

    def _create_menu(self, menus):
        # ~ https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMenu.html
        # ~ nItemId  specifies the ID of the menu item to be inserted.
        # ~ aText    specifies the label of the menu item.
        # ~ nItemStyle   0 = Standard, CHECKABLE = 1, RADIOCHECK = 2, AUTOCHECK = 4
        # ~ nItemPos specifies the position where the menu item will be inserted.
        self._menu = create_instance("com.sun.star.awt.MenuBar", True)
        for i, m in enumerate(menus):
            self._menu.insertItem(i, m["label"], m.get("style", 0), i)
            cmd = m["label"].lower().replace(" ", "_")
            self._menu.setCommand(i, cmd)
            submenu = self._create_popupmenu(m["submenu"])
            self._menu.setPopupMenu(i, submenu)

        self._window.setMenuBar(self._menu)
        return

    def _add_listeners(self, control=None):
        if self.events is None:
            return
        controller = EventsWindow(self)
        self._window.addTopWindowListener(controller)
        self._window.addWindowListener(controller)
        # ~ self._container.addKeyListener(EventsKeyWindow(self))
        return

    def _set_image_url(self, image):
        if _P.exists(image):
            return _P.to_url(image)

        path = _P.join(self._path, DIR["images"], image)
        return _P.to_url(path)

    def _special_properties(self, tipo, args):
        if tipo == "link" and not "Label" in args:
            args["Label"] = args["URL"]
            return args

        if tipo == "button":
            if "ImageURL" in args:
                args["ImageURL"] = self._set_image_url(args["ImageURL"])
            args["FocusOnClick"] = args.get("FocusOnClick", False)
            return args

        if tipo == "roadmap":
            args["Height"] = args.get("Height", self.height)
            if "Title" in args:
                args["Text"] = args.pop("Title")
            return args

        if tipo == "tree":
            args["SelectionType"] = args.get("SelectionType", SINGLE)
            return args

        if tipo == "grid":
            args["ShowRowHeader"] = args.get("ShowRowHeader", True)
            return args

        if tipo == "pages":
            args["Width"] = args.get("Width", self.width)
            args["Height"] = args.get("Height", self.height)

        return args

    def add_control(self, args):
        tipo = args.pop("Type").lower()
        root = args.pop("Root", "")
        sheets = args.pop("Sheets", ())
        columns = args.pop("Columns", ())

        args = self._special_properties(tipo, args)
        model = self.model.createInstance(self.MODELS[tipo])
        _set_properties(model, args)
        name = args["Name"]
        self.model.insertByName(name, model)
        control = self._subcont.getControl(name)
        _add_listeners(self.events, control, name)
        control = UNO_CLASSES[tipo](control)

        # ~ if tipo in ('listbox',):
        # ~ control.path = self.path

        if tipo == "tree" and root:
            control.root = root
        elif tipo == "grid" and columns:
            control.columns = columns
        elif tipo == "pages" and sheets:
            control.sheets = sheets
            control.events = self.events

        setattr(self, name, control)
        return control

    @property
    def events(self):
        return self._events

    @events.setter
    def events(self, controllers):
        self._events = controllers(self)
        self._add_listeners()

    @property
    def model(self):
        return self._model

    @property
    def width(self):
        return self._container.Size.Width

    @property
    def height(self):
        return self._container.Size.Height

    @property
    def name(self):
        return self._title.lower().replace(" ", "_")

    def add_menu(self, menus):
        self._create_menu(menus)
        return

    def open(self):
        self._window.setVisible(True)
        return

    def close(self):
        self._window.setMenuBar(None)
        self._window.dispose()
        self._frame.close(True)
        return


class LODBServer(object):
    DRIVERS = {
        "mysql": "mysqlc",
        "mariadb": "mysqlc",
        "postgres": "postgresql:postgresql",
    }
    PORTS = {
        "mysql": 3306,
        "mariadb": 3306,
        "postgres": 5432,
    }

    def __init__(self):
        self._conn = None
        self._error = "Not connected"
        self._type = ""
        self._drivers = []

    def __str__(self):
        return f"DB type {self._type}"

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.disconnet()

    @property
    def is_connected(self):
        return not self._conn is None

    @property
    def error(self):
        return self._error

    @property
    def drivers(self):
        return self._drivers

    def disconnet(self):
        if not self._conn is None:
            if not self._conn.isClosed():
                self._conn.close()
                self._conn.dispose()
        return

    def connect(self, options={}):
        args = options.copy()
        self._error = ""
        self._type = args.get("type", "postgres")
        driver = self.DRIVERS[self._type]
        server = args.get("server", "localhost")
        port = args.get("port", self.PORTS[self._type])
        dbname = args.get("dbname", "")
        user = args["user"]
        password = args["password"]

        data = {"user": user, "password": password}
        url = f"sdbc:{driver}:{server}:{port}/{dbname}"

        # ~ https://downloads.mariadb.com/Connectors/java/
        # ~ data['JavaDriverClass'] = 'org.mariadb.jdbc.Driver'
        # ~ url = f'jdbc:mysql://{server}:{port}/{dbname}'

        args = dict_to_property(data)
        manager = create_instance("com.sun.star.sdbc.DriverManager")
        self._drivers = [d.ImplementationName for d in manager]

        try:
            self._conn = manager.getConnectionWithInfo(url, args)
        except Exception as e:
            error(e)
            self._error = str(e)

        return self

    def execute(self, sql):
        query = self._conn.createStatement()
        try:
            query.execute(sql)
            result = True
        except Exception as e:
            error(e)
            self._error = str(e)
            result = False

        return result


def create_window(args):
    return LOWindow(args)


class classproperty:
    def __init__(self, method=None):
        self.fget = method

    def __get__(self, instance, cls=None):
        return self.fget(cls)

    def getter(self, method):
        self.fget = method
        return self


class ClipBoard(object):
    SERVICE = "com.sun.star.datatransfer.clipboard.SystemClipboard"
    CLIPBOARD_FORMAT_TEXT = "text/plain;charset=utf-16"

    class TextTransferable(unohelper.Base, XTransferable):
        def __init__(self, text):
            df = DataFlavor()
            df.MimeType = ClipBoard.CLIPBOARD_FORMAT_TEXT
            df.HumanPresentableName = "encoded text utf-16"
            self.flavors = (df,)
            self._data = text

        def getTransferData(self, flavor):
            return self._data

        def getTransferDataFlavors(self):
            return self.flavors

    @classmethod
    def set(cls, value):
        ts = cls.TextTransferable(value)
        sc = create_instance(cls.SERVICE)
        sc.setContents(ts, None)
        return

    @classproperty
    def contents(cls):
        df = None
        text = ""
        sc = create_instance(cls.SERVICE)
        transferable = sc.getContents()
        data = transferable.getTransferDataFlavors()
        for df in data:
            if df.MimeType == cls.CLIPBOARD_FORMAT_TEXT:
                break
        if df:
            text = transferable.getTransferData(df)
        return text


_CB = ClipBoard


class Paths(object):
    FILE_PICKER = "com.sun.star.ui.dialogs.FilePicker"
    FOLDER_PICKER = "com.sun.star.ui.dialogs.FolderPicker"

    def __init__(self, path=""):
        if path.startswith("file://"):
            path = str(Path(uno.fileUrlToSystemPath(path)).resolve())
        self._path = Path(path)

    @property
    def path(self):
        return str(self._path.parent)

    @property
    def file_name(self):
        return self._path.name

    @property
    def name(self):
        return self._path.stem

    @property
    def ext(self):
        return self._path.suffix[1:]

    @property
    def info(self):
        return self.path, self.file_name, self.name, self.ext

    @property
    def url(self):
        return self._path.as_uri()

    @property
    def size(self):
        return self._path.stat().st_size

    @classproperty
    def home(self):
        return str(Path.home())

    @classproperty
    def documents(self):
        return self.config()

    @classproperty
    def temp_dir(self):
        return tempfile.gettempdir()

    @classproperty
    def python(self):
        if IS_WIN:
            path = self.join(self.config("Module"), PYTHON)
        elif IS_MAC:
            path = self.join(self.config("Module"), "..", "Resources", PYTHON)
        else:
            path = sys.executable
        return path

    @classproperty
    def user_profile(self):
        path = self.config("UserConfig")
        path = str(Path(path).parent)
        return path

    @classproperty
    def user_config(self):
        path = self.config("UserConfig")
        return path

    @classmethod
    def dir_tmp(self, only_name=False):
        dt = tempfile.TemporaryDirectory()
        if only_name:
            dt = dt.name
        return dt

    @classmethod
    def tmp(cls, ext=""):
        tmp = tempfile.NamedTemporaryFile(suffix=ext)
        return tmp.name

    @classmethod
    def save_tmp(cls, data):
        path_tmp = cls.tmp()
        cls.save(path_tmp, data)
        return path_tmp

    @classmethod
    def config(cls, name="Work"):
        """
        Return path from config
        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XPathSettings.html
        """
        path = create_instance("com.sun.star.util.PathSettings")
        path = cls.to_system(getattr(path, name))
        return path

    @classmethod
    def get(cls, init_dir="", filters: str = ""):
        """
        Get path for save
        Options: http://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1ui_1_1dialogs_1_1TemplateDescription.html
        filters: 'xml' or 'txt,xml'
        """
        if not init_dir:
            init_dir = cls.documents
        init_dir = cls.to_url(init_dir)
        file_picker = create_instance(cls.FILE_PICKER)
        file_picker.setTitle(_("Select path"))
        file_picker.setDisplayDirectory(init_dir)
        file_picker.initialize((2,))
        if filters:
            filters = [(f.upper(), f"*.{f.lower()}") for f in filters.split(",")]
            file_picker.setCurrentFilter(filters[0][0])
            for f in filters:
                file_picker.appendFilter(f[0], f[1])

        path = ""
        if file_picker.execute():
            path = cls.to_system(file_picker.getSelectedFiles()[0])
        return path

    @classmethod
    def get_dir(cls, init_dir=""):
        folder_picker = create_instance(cls.FOLDER_PICKER)
        if not init_dir:
            init_dir = cls.documents
        init_dir = cls.to_url(init_dir)
        folder_picker.setTitle(_("Select directory"))
        folder_picker.setDisplayDirectory(init_dir)

        path = ""
        if folder_picker.execute():
            path = cls.to_system(folder_picker.getDirectory())
        return path

    @classmethod
    def get_file(cls, init_dir: str = "", filters: str = "", multiple: bool = False):
        """
        Get path file

        init_folder: folder default open
        filters: 'xml' or 'xml,txt'
        multiple: True for multiple selected
        """
        if not init_dir:
            init_dir = cls.documents
        init_dir = cls.to_url(init_dir)

        file_picker = create_instance(cls.FILE_PICKER)
        file_picker.setTitle(_("Select file"))
        file_picker.setDisplayDirectory(init_dir)
        file_picker.setMultiSelectionMode(multiple)

        if filters:
            filters = [(f.upper(), f"*.{f.lower()}") for f in filters.split(",")]
            file_picker.setCurrentFilter(filters[0][0])
            for f in filters:
                file_picker.appendFilter(f[0], f[1])

        path = ""
        if file_picker.execute():
            files = file_picker.getSelectedFiles()
            path = [cls.to_system(f) for f in files]
            if not multiple:
                path = path[0]
        return path

    @classmethod
    def replace_ext(cls, path, new_ext):
        p = Paths(path)
        name = f"{p.name}.{new_ext}"
        path = cls.join(p.path, name)
        return path

    @classmethod
    def exists(cls, path):
        result = False
        if path:
            path = cls.to_system(path)
            result = Path(path).exists()
        return result

    @classmethod
    def exists_app(cls, name_app):
        return bool(shutil.which(name_app))

    @classmethod
    def open(cls, path):
        if IS_WIN:
            os.startfile(path)
        else:
            pid = subprocess.Popen(["xdg-open", path]).pid
        return

    @classmethod
    def is_dir(cls, path):
        return Path(path).is_dir()

    @classmethod
    def is_file(cls, path):
        return Path(path).is_file()

    @classmethod
    def join(cls, *paths):
        return str(Path(paths[0]).joinpath(*paths[1:]))

    @classmethod
    def save(cls, path, data, encoding="utf-8"):
        result = bool(Path(path).write_text(data, encoding=encoding))
        return result

    @classmethod
    def save_bin(cls, path, data):
        result = bool(Path(path).write_bytes(data))
        return result

    @classmethod
    def read(cls, path, get_lines=False, encoding="utf-8"):
        if get_lines:
            with Path(path).open(encoding=encoding) as f:
                data = f.readlines()
        else:
            data = Path(path).read_text(encoding=encoding)
        return data

    @classmethod
    def read_bin(cls, path):
        data = Path(path).read_bytes()
        return data

    @classmethod
    def to_url(cls, path):
        if not path.startswith("file://"):
            path = Path(path).as_uri()
        return path

    @classmethod
    def to_system(cls, path):
        if path.startswith("file://"):
            path = str(Path(uno.fileUrlToSystemPath(path)).resolve())
        return path

    @classmethod
    def kill(cls, path):
        p = Path(path)
        try:
            if p.is_file():
                p.unlink()
            elif p.is_dir():
                shutil.rmtree(path)
            result = True
        except OSError as e:
            log.error(e)
            result = False

        return result

    @classmethod
    def files(cls, path, pattern="*"):
        files = [str(p) for p in Path(path).glob(pattern) if p.is_file()]
        return files

    @classmethod
    def dirs(cls, path):
        dirs = [str(p) for p in Path(path).iterdir() if p.is_dir()]
        return dirs

    @classmethod
    def walk(cls, path, filters=""):
        paths = []
        if filters in ("*", "*.*"):
            filters = ""
        for folder, _, files in os.walk(path):
            if filters:
                pattern = re.compile(r"\.(?:{})$".format(filters), re.IGNORECASE)
                paths += [cls.join(folder, f) for f in files if pattern.search(f)]
            else:
                paths += [cls.join(folder, f) for f in files]
        return paths

    @classmethod
    def walk_dirs(cls, path, tree=False):
        """
        Get directories recursively
        path: path source
        tree: get info in a tuple (ID_FOLDER, ID_PARENT, NAME)
        """
        folders = []
        if tree:
            i = 0
            p = 0
            parents = {path: 0}
            for root, dirs, _ in os.walk(path):
                for name in dirs:
                    i += 1
                    rn = cls.join(root, name)
                    if not rn in parents:
                        parents[rn] = i
                    folders.append((i, parents[root], name))
        else:
            for root, dirs, _ in os.walk(path):
                folders += [cls.join(root, name) for name in dirs]
        return folders

    @classmethod
    def from_id(cls, id_ext):
        pip = CTX.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        path = _P.to_system(pip.getPackageLocation(id_ext))
        return path

    @classmethod
    def from_json(cls, path):
        data = json.loads(cls.read(path))
        return data

    @classmethod
    def to_json(cls, path, data):
        data = json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True)
        return cls.save(path, data)

    @classmethod
    def from_csv(cls, path, args={}):
        # ~ See https://docs.python.org/3.7/library/csv.html#csv.reader
        with open(path) as f:
            rows = tuple(csv.reader(f, **args))
        return rows

    @classmethod
    def to_csv(cls, path, data, args={}):
        with open(path, "w") as f:
            writer = csv.writer(f, **args)
            writer.writerows(data)
        return

    @classmethod
    def zip(cls, source, target=""):
        path_zip = target
        if not isinstance(source, (tuple, list)):
            path, _, name, _ = _P(source).info
            start = len(path) + 1
            if not target:
                path_zip = f"{path}/{name}.zip"

        if isinstance(source, (tuple, list)):
            files = [(f, f[len(_P(f).path) + 1 :]) for f in source]
        elif _P.is_file(source):
            files = ((source, source[start:]),)
        else:
            files = [(f, f[start:]) for f in _P.walk(source)]

        compression = zipfile.ZIP_DEFLATED
        with zipfile.ZipFile(path_zip, "w", compression=compression) as z:
            for f in files:
                z.write(f[0], f[1])
        return path_zip

    @classmethod
    def zip_content(cls, path):
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
        return names

    @classmethod
    def unzip(cls, source, target="", members=None, pwd=None):
        path = target
        if not target:
            path = _P(source).path
        with zipfile.ZipFile(source) as z:
            if not pwd is None:
                pwd = pwd.encode()
            if isinstance(members, str):
                members = (members,)
            z.extractall(path, members=members, pwd=pwd)
        return

    @classmethod
    def merge_zip(cls, target, zips):
        try:
            with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as t:
                for path in zips:
                    with zipfile.ZipFile(path, compression=zipfile.ZIP_DEFLATED) as s:
                        for name in s.namelist():
                            t.writestr(name, s.open(name).read())
        except Exception as e:
            error(e)
            return False

        return True

    @classmethod
    def image(cls, path):
        # ~ sfa = create_instance('com.sun.star.ucb.SimpleFileAccess')
        # ~ stream = sfa.openFileRead(cls.to_url(path))
        gp = create_instance("com.sun.star.graphic.GraphicProvider")
        if isinstance(path, str):
            properties = (PropertyValue(Name="URL", Value=cls.to_url(path)),)
        else:
            properties = (PropertyValue(Name="InputStream", Value=path),)
        image = gp.queryGraphic(properties)
        return image

    @classmethod
    def copy(cls, source, target="", name=""):
        p, f, n, e = _P(source).info
        if target:
            p = target
        e = f".{e}"
        if name:
            e = ""
            n = name
        path_new = cls.join(p, f"{n}{e}")
        shutil.copy(source, path_new)
        return path_new


_P = Paths


class Dates(object):
    @classmethod
    def date(cls, year, month, day):
        d = datetime.date(year, month, day)
        return d

    @classmethod
    def str_to_date(cls, str_date, template, to_calc=False):
        d = datetime.datetime.strptime(str_date, template).date()
        if to_calc:
            d = d.toordinal() - DATE_OFFSET
        return d

    @classmethod
    def calc_to_date(cls, value, frm=""):
        d = datetime.date.fromordinal(int(value) + DATE_OFFSET)
        if frm:
            d = d.strftime(frm)
        return d


class OutputStream(unohelper.Base, XOutputStream):
    def __init__(self):
        self._buffer = b""
        self.closed = 0

    @property
    def buffer(self):
        return self._buffer

    def closeOutput(self):
        self.closed = 1

    def writeBytes(self, seq):
        if seq.value:
            self._buffer = seq.value

    def flush(self):
        pass


class IOStream(object):
    @classmethod
    def buffer(cls):
        return io.BytesIO()

    @classmethod
    def input(cls, buffer):
        instance = "com.sun.star.io.SequenceInputStream"
        stream = create_instance(instance, True)
        stream.initialize((uno.ByteSequence(buffer.getvalue()),))
        return stream

    @classmethod
    def output(cls):
        return OutputStream()

    @classmethod
    def qr(cls, data, **kwargs):
        import segno

        kwargs["kind"] = kwargs.get("kind", "svg")
        kwargs["scale"] = kwargs.get("scale", 8)
        kwargs["border"] = kwargs.get("border", 2)
        buffer = cls.buffer()
        segno.make(data).save(buffer, **kwargs)
        stream = cls.input(buffer)
        return stream


class SpellChecker(object):
    def __init__(self):
        service = "com.sun.star.linguistic2.SpellChecker"
        self._spellchecker = create_instance(service, True)
        self._locale = LOCALE

    @property
    def locale(self):
        slocal = f"{self._locale.Language}-{self._locale.Country}"
        return slocale

    @locale.setter
    def locale(self, value):
        lang = value.split("-")
        self._locale = Locale(lang[0], lang[1], "")

    def is_valid(self, word):
        result = self._spellchecker.isValid(word, self._locale, ())
        return result

    def spell(self, word):
        result = self._spellchecker.spell(word, self._locale, ())
        if result:
            result = result.getAlternatives()
            if not isinstance(result, tuple):
                result = ()
        return result


def spell(word, locale=""):
    sc = SpellChecker()
    if locale:
        sc.locale = locale
    return sc.spell(word)


def __getattr__(name):
    if name == "active":
        return LODocs().active
    if name == "active_sheet":
        return LODocs().active.active
    if name == "selection":
        return LODocs().active.selection
    if name == "current_region":
        return LODocs().active.selection.current_region
    if name in ("rectangle", "pos_size"):
        return Rectangle()
    if name == "paths":
        return Paths
    if name == "docs":
        return LODocs()
    if name == "db":
        return LODBServer()
    if name == "sheets":
        return LOSheets()
    if name == "cells":
        return LOCells()
    if name == "menus":
        return LOMenus()
    if name == "shortcuts":
        return LOShortCuts()
    if name == "clipboard":
        return ClipBoard
    if name == "dates":
        return Dates
    if name == "ios":
        return IOStream()
    raise AttributeError(f"module '{__name__}' has no attribute '{name}'")


def create_dialog(args):
    return LODialog(args)


def inputbox(message, default="", title=TITLE, echochar=""):
    class ControllersInput(object):
        def __init__(self, dlg):
            self.d = dlg

        def cmd_ok_action(self, event):
            self.d.close(1)
            return

    args = {
        "Title": title,
        "Width": 200,
        "Height": 80,
    }
    dlg = LODialog(args)
    dlg.events = ControllersInput

    args = {
        "Type": "Label",
        "Name": "lbl_msg",
        "Label": message,
        "Width": 140,
        "Height": 50,
        "X": 5,
        "Y": 5,
        "MultiLine": True,
        "Border": 1,
    }
    dlg.add_control(args)

    args = {
        "Type": "Text",
        "Name": "txt_value",
        "Text": default,
        "Width": 190,
        "Height": 15,
    }
    if echochar:
        args["EchoChar"] = ord(echochar[0])
    dlg.add_control(args)
    dlg.txt_value.move(dlg.lbl_msg)

    args = {
        "Type": "button",
        "Name": "cmd_ok",
        "Label": _("OK"),
        "Width": 40,
        "Height": 15,
        "DefaultButton": True,
        "PushButtonType": 1,
    }
    dlg.add_control(args)
    dlg.cmd_ok.move(dlg.lbl_msg, 10, 0)

    args = {
        "Type": "button",
        "Name": "cmd_cancel",
        "Label": _("Cancel"),
        "Width": 40,
        "Height": 15,
        "PushButtonType": 2,
    }
    dlg.add_control(args)
    dlg.cmd_cancel.move(dlg.cmd_ok)

    if dlg.open():
        return dlg.txt_value.value

    return ""


def get_fonts():
    toolkit = create_instance("com.sun.star.awt.Toolkit")
    device = toolkit.createScreenCompatibleDevice(0, 0)
    return device.FontDescriptors


def get_filters():
    """
    Get all support filters
    https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
    """
    factory = create_instance("com.sun.star.document.FilterFactory")
    rows = [data_to_dict(factory[name]) for name in factory]
    for row in rows:
        row["UINames"] = data_to_dict(row["UINames"])
    return rows


# ~ https://en.wikipedia.org/wiki/Web_colors
def get_color(value):
    COLORS = {
        "aliceblue": 15792383,
        "antiquewhite": 16444375,
        "aqua": 65535,
        "aquamarine": 8388564,
        "azure": 15794175,
        "beige": 16119260,
        "bisque": 16770244,
        "black": 0,
        "blanchedalmond": 16772045,
        "blue": 255,
        "blueviolet": 9055202,
        "brown": 10824234,
        "burlywood": 14596231,
        "cadetblue": 6266528,
        "chartreuse": 8388352,
        "chocolate": 13789470,
        "coral": 16744272,
        "cornflowerblue": 6591981,
        "cornsilk": 16775388,
        "crimson": 14423100,
        "cyan": 65535,
        "darkblue": 139,
        "darkcyan": 35723,
        "darkgoldenrod": 12092939,
        "darkgray": 11119017,
        "darkgreen": 25600,
        "darkgrey": 11119017,
        "darkkhaki": 12433259,
        "darkmagenta": 9109643,
        "darkolivegreen": 5597999,
        "darkorange": 16747520,
        "darkorchid": 10040012,
        "darkred": 9109504,
        "darksalmon": 15308410,
        "darkseagreen": 9419919,
        "darkslateblue": 4734347,
        "darkslategray": 3100495,
        "darkslategrey": 3100495,
        "darkturquoise": 52945,
        "darkviolet": 9699539,
        "deeppink": 16716947,
        "deepskyblue": 49151,
        "dimgray": 6908265,
        "dimgrey": 6908265,
        "dodgerblue": 2003199,
        "firebrick": 11674146,
        "floralwhite": 16775920,
        "forestgreen": 2263842,
        "fuchsia": 16711935,
        "gainsboro": 14474460,
        "ghostwhite": 16316671,
        "gold": 16766720,
        "goldenrod": 14329120,
        "gray": 8421504,
        "grey": 8421504,
        "green": 32768,
        "greenyellow": 11403055,
        "honeydew": 15794160,
        "hotpink": 16738740,
        "indianred": 13458524,
        "indigo": 4915330,
        "ivory": 16777200,
        "khaki": 15787660,
        "lavender": 15132410,
        "lavenderblush": 16773365,
        "lawngreen": 8190976,
        "lemonchiffon": 16775885,
        "lightblue": 11393254,
        "lightcoral": 15761536,
        "lightcyan": 14745599,
        "lightgoldenrodyellow": 16448210,
        "lightgray": 13882323,
        "lightgreen": 9498256,
        "lightgrey": 13882323,
        "lightpink": 16758465,
        "lightsalmon": 16752762,
        "lightseagreen": 2142890,
        "lightskyblue": 8900346,
        "lightslategray": 7833753,
        "lightslategrey": 7833753,
        "lightsteelblue": 11584734,
        "lightyellow": 16777184,
        "lime": 65280,
        "limegreen": 3329330,
        "linen": 16445670,
        "magenta": 16711935,
        "maroon": 8388608,
        "mediumaquamarine": 6737322,
        "mediumblue": 205,
        "mediumorchid": 12211667,
        "mediumpurple": 9662683,
        "mediumseagreen": 3978097,
        "mediumslateblue": 8087790,
        "mediumspringgreen": 64154,
        "mediumturquoise": 4772300,
        "mediumvioletred": 13047173,
        "midnightblue": 1644912,
        "mintcream": 16121850,
        "mistyrose": 16770273,
        "moccasin": 16770229,
        "navajowhite": 16768685,
        "navy": 128,
        "oldlace": 16643558,
        "olive": 8421376,
        "olivedrab": 7048739,
        "orange": 16753920,
        "orangered": 16729344,
        "orchid": 14315734,
        "palegoldenrod": 15657130,
        "palegreen": 10025880,
        "paleturquoise": 11529966,
        "palevioletred": 14381203,
        "papayawhip": 16773077,
        "peachpuff": 16767673,
        "peru": 13468991,
        "pink": 16761035,
        "plum": 14524637,
        "powderblue": 11591910,
        "purple": 8388736,
        "red": 16711680,
        "rosybrown": 12357519,
        "royalblue": 4286945,
        "saddlebrown": 9127187,
        "salmon": 16416882,
        "sandybrown": 16032864,
        "seagreen": 3050327,
        "seashell": 16774638,
        "sienna": 10506797,
        "silver": 12632256,
        "skyblue": 8900331,
        "slateblue": 6970061,
        "slategray": 7372944,
        "slategrey": 7372944,
        "snow": 16775930,
        "springgreen": 65407,
        "steelblue": 4620980,
        "tan": 13808780,
        "teal": 32896,
        "thistle": 14204888,
        "tomato": 16737095,
        "turquoise": 4251856,
        "violet": 15631086,
        "wheat": 16113331,
        "white": 16777215,
        "whitesmoke": 16119285,
        "yellow": 16776960,
        "yellowgreen": 10145074,
    }

    if isinstance(value, tuple):
        color = (value[0] << 16) + (value[1] << 8) + value[2]
    else:
        if value[0] == "#":
            r, g, b = bytes.fromhex(value[1:])
            color = (r << 16) + (g << 8) + b
        else:
            color = COLORS.get(value.lower(), -1)
    return color


COLOR_ON_FOCUS = get_color("LightYellow")


class LOServer(object):
    HOST = "localhost"
    PORT = "8100"
    ARG = f"socket,host={HOST},port={PORT};urp;StarOffice.ComponentContext"
    CMD = [
        "soffice",
        "-env:SingleAppInstance=false",
        "-env:UserInstallation=file:///tmp/LO_Process8100",
        "--headless",
        "--norestore",
        "--invisible",
        f"--accept={ARG}",
    ]

    def __init__(self):
        self._server = None
        self._ctx = None
        self._sm = None
        self._start_server()
        self._init_values()

    def _init_values(self):
        global CTX
        global SM

        if not self.is_running:
            return

        ctx = uno.getComponentContext()
        service = "com.sun.star.bridge.UnoUrlResolver"
        resolver = ctx.ServiceManager.createInstanceWithContext(service, ctx)
        self._ctx = resolver.resolve("uno:{}".format(self.ARG))
        self._sm = self._ctx.getServiceManager()
        CTX = self._ctx
        SM = self._sm
        return

    @property
    def is_running(self):
        try:
            s = socket.create_connection((self.HOST, self.PORT), 5.0)
            s.close()
            debug("LibreOffice is running...")
            return True
        except ConnectionRefusedError:
            return False

    def _start_server(self):
        if self.is_running:
            return

        for i in range(3):
            self._server = subprocess.Popen(self.CMD, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            time.sleep(3)
            if self.is_running:
                break
        return

    def stop(self):
        if self._server is None:
            print("Search pgrep soffice")
        else:
            self._server.terminate()
        debug("LibreOffice is stop...")
        return

    def create_instance(self, name, with_context=True):
        if with_context:
            instance = self._sm.createInstanceWithContext(name, self._ctx)
        else:
            instance = self._sm.createInstance(name)
        return instance
