#!/usr/bin/python
# coding=utf-8'
#
# Copyright (c)  Tomasz Sługocki ts.kontakt@gmail.com
# This code is licensed under Apache 2.0
from math import log
import os
import subprocess
import sys

import xlsxwriter as xls

FONT_SIZE = 9
HEADER_FONT_NAME = "Arial"
TITLE_RANGE = "B1:E1"
HEADER_BG_COLOR = "#D4D0C8"
HEADER_FONT_COLOR = "#003366"


def calculate_column_width(text_length):
    # Adjust column width based on header  text length
    extra_width = text_length / 10 + 10 if text_length > 20 else 0

    if text_length <= 3:
        return 3
    else:
        return (((51.16 * log(text_length) - 38.395) * 0.85) / 10 + (log(text_length) * 2) +
                extra_width)


def ensure_unicode(input_value):
    if "bytes" in repr(type(input_value)):
        return input_value.decode("utf8")
    return input_value


def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


class Exfile(object):

    def __init__(self, filename):
        assert ".xlsx" in filename
        filename.upper()
        self.workbook = xls.Workbook(filename)

        file_path = (filename if os.sep in filename else os.path.join(os.getcwd(), filename))
        self.file_path = file_path

    def write(self, data_list, title, worksheet_name=None, wrap=False):
        assert (isinstance(data_list[0], list) or isinstance(data_list[0], tuple) or
                "numpy." in repr(type(data_list[0])) or "dict_values" in repr(data_list[0]))

        title.upper()
        if worksheet_name:
            worksheet_name.upper()
        # global configs

        worksheet = self.workbook.add_worksheet(ensure_unicode(worksheet_name))
        number_format = "#,##0;[Red]-#,##0"
        start_row = 4
        start_column = 1

        merge_format = self.workbook.add_format({
            "bold": 0,
            "border": 1,
            "align": "center",
            "fg_color": "#F1F1F1",
            "text_wrap": True,
            "font_name": HEADER_FONT_NAME,
            "font_size": FONT_SIZE,
        })

        worksheet.merge_range(TITLE_RANGE, ensure_unicode(title), merge_format)

        header_format = self.workbook.add_format()
        header_format.set_bold()
        header_format.set_font_color(HEADER_FONT_COLOR)
        header_format.set_font_size(9)
        header_format.set_bg_color(HEADER_BG_COLOR)
        header_format.set_indent(1)

        cell_format = self.workbook.add_format({
            "font_name": "Arial",
            "font_size": FONT_SIZE,
            "num_format": number_format,
        })

        decimal_format = self.workbook.add_format({
            "font_size": FONT_SIZE,
            "num_format": "0.00;[RED]-0.00"
        })

        long_text_format = self.workbook.add_format({"text_wrap": True, "font_size": 8})
        long_text_format.set_align("vcenter")

        for column_index, header_value in enumerate(data_list[0]):
            text_length = len(repr(header_value))
            column_width = calculate_column_width(text_length)
            worksheet.set_column(start_column + column_index, start_column + column_index,
                                 column_width)

        worksheet.freeze_panes(start_row, 0)

        for column_index, header_value in enumerate(data_list[0]):
            if "str" in repr(type(header_value)):
                header_value = ensure_unicode(header_value)

            worksheet.write(
                start_row - 1,
                start_column + column_index,
                str(header_value).strip(),
                header_format,
            )

        for row_index, row_data in enumerate(data_list[1:]):
            current_row = start_row + row_index
            for column_index, cell_value in enumerate(row_data):
                if "str" in repr(type(cell_value)):
                    try:
                        cell_value = ensure_unicode(cell_value)
                    except UnicodeDecodeError:
                        print("! decode error")
                        print([cell_value])

                    if wrap:
                        worksheet.write(
                            current_row,
                            start_column + column_index,
                            cell_value,
                            long_text_format,
                        )
                    else:
                        worksheet.write(
                            current_row,
                            start_column + column_index,
                            cell_value,
                            cell_format,
                        )
                else:
                    try:
                        if abs(cell_value) < 20.0 and int(cell_value) != round(cell_value, 2):
                            worksheet.write(
                                current_row,
                                start_column + column_index,
                                cell_value,
                                decimal_format,
                            )
                        else:
                            worksheet.write(
                                current_row,
                                start_column + column_index,
                                cell_value,
                                cell_format,
                            )
                    except TypeError:
                        cell_value = " " + str(repr(cell_value))[:200]
                        worksheet.write(
                            current_row,
                            start_column + column_index,
                            cell_value,
                            cell_format,
                        )
        # worksheet.autofit() this is not working well
        return worksheet

    def add_links(self):
        start_column = 6
        start_row = 0
        worksheets = self.workbook.worksheets()

        inactive_format = self.workbook.add_format({
            "font_color": "gray",
            "bold": 0,
            "underline": 1,
            # "font_size": 10,
        })

        for source_sheet in worksheets:
            link_index = 0
            for target_sheet in worksheets:
                if target_sheet.name == source_sheet.name:
                    source_sheet.write_url(
                        start_row,
                        start_column + link_index,
                        "internal:'%s'!A1" % target_sheet.name,
                        inactive_format,
                        string=target_sheet.name,
                    )
                else:
                    source_sheet.write_url(
                        start_row,
                        start_column + link_index,
                        "internal:'%s'!A1" % target_sheet.name,
                        string=target_sheet.name,
                    )
                link_index += 1

    def save(self):
        try:
            self.workbook.close()
        except BaseException:
            print(sys.exc_info())
            if "denied" in repr(sys.exc_info()):
                print("! Cannot write file - permission error")
                self.workbook.close()
        finally:
            pass


def save_list(xls_name, inlist, header_list=None, title="Title", shname="sheet1", wrap=False):
    exfile = Exfile(xls_name)
    if header_list:
        assert header_list.extend
        if hasattr(inlist, 'to_records'):
            print(type(inlist))
            inlist = list(inlist.to_records())
        if hasattr(inlist, 'tolist'):
            inlist = inlist.tolist()

        inlist.insert(0, header_list)
    exfile.write(inlist, title, shname, wrap=wrap)
    try:
        exfile.save()
        open_file(xls_name)
    except BaseException:
        if "permission" in repr(sys.exc_info()).lower():
            # works on windows only, requires nicmd utility
            cmd = 'nircmd win close stitle "{xls_name}"'
            (stdout, strerr) = subprocess.Popen(cmd, shell=True,
                                                stdout=subprocess.PIPE).communicate()
            print("!--", stdout)
            print("---", strerr)
            # close_file(xls_name)
            exfile.save()
            open_file(xls_name)


def generate_random_data(num_rows=10):
    import random
    import string
    from datetime import datetime, timedelta

    unicode_ranges = [
        (0x0020, 0x007E),  # Basic Latin (printable ASCII)
        (0x00A0, 0x00FF),  # Latin-1 Supplement (e.g., accented characters)
        (0x0100, 0x017F),  # Latin Extended-A (more European characters)
        # (0x0370, 0x03FF),  # Greek and Coptic
    ]

    def gen_datetime(min_year=1900, max_year=datetime.now().year):
        # generate a datetime in format yyyy-mm-dd hh:mm:ss.000000
        start = datetime(min_year, 1, 1, 00, 00, 00)
        years = max_year - min_year + 1
        end = start + timedelta(days=365 * years)
        return start + (end - start) * random.random()

    def get_random_unicode_char():
        """Get a random Unicode character from various language ranges."""
        range_start, range_end = random.choice(unicode_ranges)
        code_point = random.randint(range_start, range_end)
        try:
            return chr(code_point)
        except ValueError:
            return chr(random.randint(0x00C0, 0x00FF))

    header = ['col1', 'col2', 'col3', 'col4', 'col5', 'col6']

    result = [header]
    for _ in range(num_rows):
        row = [
            random.choice(string.ascii_letters),
            random.randint(100, 100000),
            random.uniform(-1, 1),
            get_random_unicode_char(),
            random.choice([True, False]),
            str(gen_datetime()),
        ]
        result.append(row)

    return result


def get_packages():
    try:
        import pkg_resources

        dists = [repr(d).split(" ") for d in sorted(pkg_resources.working_set)]
        dists = sorted(dists, key=lambda x: x[0].lower())
        header_list = ["name        ", "ver  ", "full package path"]
        dists.insert(0, header_list)
    except ModuleNotFoundError:
        print(f"Error loading module pkg_resources - using random data")
        dists = generate_random_data(num_rows=100)
        dists.insert(0, ["name1", "name2", "name3", "name4", "name5", "name6"])
    return dists


def test_colwidth():
    import pandas as pd
    import re

    # print(pd._build_option_description)
    a = pd.describe_option(_print_desc=False)

    outlist = []
    for line in a.split("\n"):
        if re.search("^[a-z]", line):
            data = [line, " -> "]
        else:
            data = [" ", line]
        outlist.append(data)
    header = ["Option             ", "Description           "]
    save_list("test.xlsx", outlist, header_list=header, title="All pandas registered options")


def test_simple():
    save_list("test.xlsx", get_packages(), title="Current user python packages")


def test_numpy():
    import numpy as np
    from numpy.random import default_rng
    arr = default_rng(42).random((100, 4))

    # import pandas as pd

    # df  = pd.DataFrame(arr)

    header = ['col1', 'col2', 'col3', 'col4']
    save_list("test.xlsx", arr, header, title="Test numpy")


def system_info():
    import platform
    import re
    import socket

    try:
        info = {}
        info["platform"] = platform.system()
        info["platform-release"] = platform.release()
        info["platform-version"] = platform.version()
        info["architecture"] = platform.machine()
        info["hostname"] = socket.gethostname()
        info["ip-address"] = socket.gethostbyname(socket.gethostname())
        info["mac-address"] = ":".join(re.findall("..", "%012x" % uuid.getnode()))
        info["processor"] = platform.processor()
        info["ram"] = (str(round(psutil.virtual_memory().total / (1024.0**3))) + " GB")
    except Exception as e:
        print(sys.exc_info())
    return info


def test_multisheets():

    #some example data
    data1 = get_packages()
    data2 = generate_random_data(20)
    data3 = [(x, y) for x, y in system_info().items()]
    #create file
    exfile = Exfile("test_multisheet_file.xlsx")
    exfile.write(data1, title="Current user packages")
    exfile.write(data2, title="Random data")
    exfile.write(data3, title="System Info", wrap=True)
    exfile.add_links()
    exfile.save()
    open_file("test_multisheet_file.xlsx")


if __name__ == "__main__":
    # test_numpy()
    # stop
    # test_simple()
    # test_colwidth()
    test_multisheets()
