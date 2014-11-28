#!/usr/bin/python
#coding=utf-8
import os
import sys
import cPickle
import pickle

def get_files(path, exts=None):
    def check_files(file, exts):
        f_ext = os.path.splitext(file)[1]
        if f_ext in exts:
            return True
        return False

    ABS_PATH = os.path.abspath(path)
    file_list = [ABS_PATH + os.sep + file for file in os.listdir(path)
                 if os.path.isfile(ABS_PATH + os.sep + file)]
    if not exts:
        return file_list
    else:
        return [file for file in file_list if check_files(file, exts)]

def get_filename(file_path):
    return os.path.split(os.path.splitext(file_path)[0])[-1] 

def integer_or_str(elem):
    result = True
    try:
        assert_num = elem / 1.
        if elem != int(elem):
            result = False
            print "! only integers allowed"
    except TypeError:
        try:
            assert_str = elem + "str"
        except:
            result = False
    return result
    
 
def just_load(file_name, mode="rb"):
    if not os.path.isfile(file_name):
        print "file %s does not exist" % file_name
        return False
    try:
        py_obj = cPickle.load(open(file_name, mode))
    except:
        print "Cpickle erro file:%s  %s\n---" % (file_name,
                repr(sys.exc_info()))
        try:
            #try load with different mode
            py_obj = pickle.load(open(file_name, "r"))
        except:
            print sys.exc_info()
            return None
    return py_obj


def just_save(file_name, py_obj, std_pickle=True, mode="wb"):
    assert_str = file_name + "str"
    
    file = open(file_name, mode)
    if std_pickle:
        pickle.dump(py_obj, file)
    else:
        cPickle.dump(py_obj, file)
    file.close()
    return
    
def save_with_key(file_name, py_obj, key):    
    if not integer_or_str(key):
        raise Exception("only integers or strigns allowed as key")
        
    assert not integer_or_str(py_obj)
    
    out_obj = { "_key" : key, "_data" : py_obj }
    just_save(file_name, out_obj)
    return
    
    
def load_with_key(file_name, key, verbose=0):
    loaded_obj = just_load(file_name, mode="rb")
    result = False
    message = ""
    if not loaded_obj:
        message = "cache [%s] file does not exist" % file_name
    else:
        file_key = loaded_obj.get("_key")
        if file_key:
            if file_key == key:
                message = "OK valid file '%s' loaded.\n" % file_name
                result = loaded_obj["_data"]
            else:
                message = "keys does not match %s != %s" %(file_key , key)
        else:
            message = "- not a valid cache file '_key' property not found"
    if verbose:
        print message
    return result
                

def go_up(level, file=None):
    if not file:
        up_path = os.path.split(os.getcwd())
    else:
        level += 1
        up_path = os.path.split(os.path.abspath(file))
    if level == 1:
        return up_path[0]
    new_path = None
    for x in range(level - 1):
        if new_path:
            path = os.path.split(new_path[0])
        else:
            path = os.path.split(up_path[0])
        new_path = path
    return os.path.abspath(path[0])


def get_dirname(file):
    assert os.path.isfile(file)
    return os.path.split(os.path.abspath(file))[0] + os.sep
    

def save_textfile(path, instr):
    assert_both_str =  "str".join((path, instr))
    del assert_both_str
    #~ print os.path.exists(path)
    if len(path) > 150:
        print "! path seems to be to big"
    outfile = open(path, 'w')
    outfile.write(instr)
    outfile.close()
    
    

def compact_float(num):
    """
    http://stackoverflow.com/questions/2440692/formatting-floats-in-python-without-superfluous-zeros
    """
    return ('%.2f' % num).rstrip('0').rstrip('.') 


__all__ = [compact_float, save_with_key, save_textfile]

if __name__ == "__main__":
    #~ import getpage
    #~ tmpurl = 'http://www.ceneo.pl`/GetShopNIP?guid=bd9bc19e-9789-45ce-901c-706fd9c32ff2'
    #~ print getpage.gethtml(tmpurl)
    #~ stop

    print 1
    import time
    for i in range(1, 5):
        print i
        time.sleep(4)
    try:
        1/0
    except Exception, e:
        print 1
        print str(e)
        print sys.exc_info()[1]

    import string 
    print dsdaad
    print 'aaaa'
    print save_textfile("vtestfile", "instri")
    
    save_with_key("mytestfile", {2131: "@@@"}, 4)
    print load_with_key("mytestfile", 1, verbose=1)
    print go_up(1, __file__)
    
    #~ print just_load("C:\\pycache\\20121023_ta_db")
    print just_load("E:\\0_Uptrend\\RES\\htmlgen\\web_modules\\bhist.cp")

    
    
    print ""

