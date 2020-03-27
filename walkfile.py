# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import os
import time


def timestamp2time(timestamp):
    """Convert time to structured format"""

    time_stamp = time.localtime(timestamp)
    time_str = time.strftime('%Y-%m-%d %H:%M:%S', time_stamp)
    return time_str


def walk_file(file_dirs):
    """
    Walk through the selected file_dirs and fetch the files’info as a
    dictionary(Key:Value). The Keys:'Name':file name; 'Extension': file format;
    'Ctime': created time;'Mtime': last modified time; 'Path': absolute path of
    file; 'Size': the size of file(Unit:Mb); 'Text': extracted text(empty now).

    The parameter file_dirs should be a list, e.g. ["C:/tempt", "D:/test"],
    which is transmitted based on the user's selections.
    """

    files_info = []
    for f_dir in file_dirs:  # if chose more than 1 file_dir
        for root, dirs, files in os.walk(f_dir):
            for f in files:
                path = os.path.join(root, f)
                extension = os.path.splitext(f)[1]
                name = os.path.splitext(f)[0]
                ctime = timestamp2time(os.path.getctime(path))
                mtime = timestamp2time(os.path.getmtime(path))
                size = os.path.getsize(path) / float(1024 * 1024)  # Mb
                info = {'Name': name, 'Extension': extension,
                        'Ctime': ctime, 'Mtime': mtime, 'Path': path,
                        'Size': size, 'Text': ''}
                files_info.append(info)

    return files_info
