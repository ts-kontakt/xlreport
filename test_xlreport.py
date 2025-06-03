#!/usr/bin/python
# coding=utf-8'
#
# Copyright (c)  Tomasz Sługocki ts.kontakt@gmail.com
# This code is licensed under Apache 2.0
import xlreport as xl
import re
import random
import string
import sys 
from datetime import datetime, timedelta

def generate_random_data(num_rows=10):

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


def get_pandas_opts():
    import pandas as pd
    options_str = pd.describe_option(_print_desc=False)
    header = ["Option             ", "Description              "]
    outlist = [header]
    for line in options_str.split("\n"):
        if re.search("^[a-z]", line):
            data = [line, " -> "]
        else:
            data = [" ", line]
        outlist.append(data)
    return outlist


def test_colwidth():
    opt_list = get_pandas_opts()
    xl.to_file("test.xlsx", opt_list, title="All pandas registered options")


def test_simple():
    xl.to_file("test.xlsx", get_packages(), title="Current user python packages")


def test_numpy():
    from numpy.random import default_rng
    arr = default_rng(42).random((100, 4))
    header = ['col1', 'col2', 'col3', 'col4']
    xl.to_file("test.xlsx", arr, header, title="Test numpy")


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
        pass
        # print(sys.exc_info())
    return info





def test_multisheets():
    # some example data
    data1 = get_pandas_opts()
    data2 = generate_random_data(20)
    stop
    data3 = [(x, y) for x, y in system_info().items()]
    # create file
    exfile = xl.Exfile("test_multisheet_file.xlsx")
    exfile.write(data1, title="Current user packages")
    exfile.write(data2, title="Random data")
    exfile.write(data3, title="System Info", wrap=True)
    exfile.add_links()
    exfile.save(start=True)



if __name__ == "__main__":
    test_simple()
    test_multisheets()
    # test_numpy()
    # test_colwidth()
    
