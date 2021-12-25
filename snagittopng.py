# Converts the SNAGIT raw capture to PNG
# this does not copy the SNAG IT annotations
# use SNAGIT Batch conversion tool to convert with Annotations


import os
import time
import datetime
from ctypes import create_unicode_buffer, windll, wintypes, byref
from bitstring import ConstBitStream


# get the file creation date/time
def get_ctime(filepath, verbose="n"):
    if verbose == "y":
        print ("filepath: " + filepath)
    created = os.path.getctime(filepath)
    return created


# get the file last modified date/time
def get_mtime(filepath, verbose="n"):
    if verbose == "y":
        print ("filepath: " + filepath)

    modified = os.path.getmtime(filepath)
    return modified


# update file creation date/time
def update_ctime(filepath, epoch, verbose="n"):
    if verbose == "y":
        print (filepath + " updating create time to: " + epoch)

    # Convert Unix timestamp to Windows FileTime using some magic numbers
    # See documentation: https://support.microsoft.com/en-us/help/167296
    timestamp = int((epoch * 10000000) + 116444736000000000)
    ctime = wintypes.FILETIME(timestamp & 0xFFFFFFFF, timestamp >> 32)

    # Call Win32 API to modify the file creation date
    handle = windll.kernel32.CreateFileW(filepath, 256, 0, None, 3, 128, None)
    windll.kernel32.SetFileTime(handle, byref(ctime), None, None)
    windll.kernel32.CloseHandle(handle)

    if verbose == "y":
        print("file create date updated")


# update file last modified date/time
def update_mtime(filepath, epoch, verbose="n"):
    if verbose == "y":
        print (filepath + " updating last modified time to: " + epoch)

    os.utime(filepath, (epoch, epoch))

    if verbose == "y":
        print("file last modified date updated")


# SNAGIT Screenshot folder
Local_App_Data = os.environ['LOCALAPPDATA']
directory = Local_App_Data+r"\TechSmith\Snagit\DataStore" + "\\"

# Current Folder
# directory = ".\\"

for filename in os.listdir(directory):
    if filename.endswith(".SNAG"):
        filepath = directory + filename

        file_create_date = get_ctime(filepath)
        file_modified_date = get_mtime(filepath)

        print("Processing   %s  %d %d" % (filepath,
                                          file_create_date,
                                          file_modified_date))

        # initialise bitstream using filename
        # it Can be initialise from files, bytes, etc.
        s = ConstBitStream(filename=filepath)

        # look for PNG Start and End Markers in SNAGIT File
        # Start marker  ‰PNG
        # end marker  END®B`‚
        start = s.find('0x89504E47', bytealigned=True)
        end = s.find('0x454E44AE426082', bytealigned=True)

        assert start, "PNG Header not found"
        assert end,  "PNG Terminator not found"

        start_off = int(start[0] / 8)
        end_off = int(end[0] / 8)
        # calculate content length, add 7 bytes to end-offset
        # as it is starting location of PNG end marker
        content_len = end_off - start_off + 7

        # Now with correct start offset and content length
        # read the file in binary mode and copy the content
        # to new file

        # File read operation start
        file = open(filepath, "rb")
        file.seek(start_off, 0)
        pngdata = file.read(content_len)
        file.close()
        # File read operation end

        # write the content to new file with png extension
        filepath_dest = filepath + ".png"
        destfile = open(filepath_dest, "wb")
        destfile.write(pngdata)
        destfile.close()

        # update the file create and last modified date time
        # to match original file
        update_ctime(filepath_dest, file_create_date)
        update_mtime(filepath_dest, file_modified_date)

    else:
        continue
