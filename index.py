# _*_ coding: UTF-8 _*_
# @version: python 3.7.4
# @author: frank0713


import os
import time
import uncompress
import extracttext


def timestamp2time(timestamp):
    """Convert time to structured format"""

    time_stamp = time.localtime(timestamp)
    time_str = time.strftime('%Y-%m-%d %H:%M:%S', time_stamp)
    return time_str


def rewrite_early(index_dir, contents):
    """
    Rewrite the earliest modified file. Save the indexing files'information, for
    comparing when renew the index.
    :param index_dir: A string. Directory - storing the temporal indexing
    files'information.
    :param contents: A list, consists of dictionaries. Acturally, It is the
    result returned by Function walk_file().
    """

    all_files = os.listdir(index_dir)
    all_files.sort(key=lambda f: os.path.getmtime(index_dir + "\\" + f))
    early_file = os.path.join(index_dir, all_files[0])
    with open(early_file, "w") as ef:
        for i in contents:
            ef.write(str(i))
            ef.write('\n')


def read_new(index_dir):
    """
    Read the lastest modified file. Read the indexing files' information for
    extracting text.
    :param index_dir: A string. Directory - storing the temporal indexing
    files'information.
    :return: A list, consists of dictionaries. Acturally, It is the
    result returned by Function walk_file() and the contents writen by Function
    rewrite_early().
    """

    all_files = os.listdir(index_dir)
    all_files.sort(key=lambda f: os.path.getmtime(index_dir + "\\" + f))
    new_file = os.path.join(index_dir, all_files[-1])
    contents = []
    with open(new_file, "r") as nf:
        content = nf.readlines()
        for line in content:
            info = line.strip('\n')
            contents.append(eval(info))
    return contents


def get_uncompress_text():
    """
    Extract the text and zipped files' name.

    :return: uncompress_info: A list, consists of dictionaries. For each
    dictionary, the Keys:'Name':file name with format; 'Text': extracted text.
    """

    uncompress_info = []
    # ToDo: 文本过滤器
    scripts = ['.py', '.r', '.cpp']
    for root, dirs, files in os.walk("C:\\temp_zip"):
        for f in files:
            ext = os.path.splitext(f)[1]
            p = os.path.join(root, f)
            name = f
            # Extract text
            if ext == '.txt' or ext == '.md' or ext == '.yml':
                text = extracttext.TXTText(p).txttext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext in scripts:
                text = extracttext.TXTText(p).txttext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.csv':
                text = extracttext.TXTText(p).csvtext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.xls' or ext == '.xlsx' or ext == '.xlsm':
                text = extracttext.TXTText(p).exceltext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.docx':
                text = extracttext.WordText(p).docxtext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.doc' or ext == '.docm' or ext == '.rtf':
                text = extracttext.WordText(p).doctext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.pptx':
                text = extracttext.PPTText(p).pptxtext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.ppt' or ext == '.pptm':
                text = extracttext.PPTText(p).ppttext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.odt' or ext == '.ods' or ext == '.odp':
                text = extracttext.ODFText(p).odftext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.xml':
                text = extracttext.MarkupText(p).xmltext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.html':
                text = extracttext.MarkupText(p).htmltext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.tex':
                text = extracttext.MarkupText(p).textext()
                info = {'Name': name, 'Text': text}
                uncompress_info.append(info)
            if ext == '.pdf':
                text = extracttext.PDFText(p).docpdftext()
                if text != '':
                    text = text
                    info = {'Name': name, 'Text': text}
                    uncompress_info.append(info)
                else:
                    text = extracttext.PDFText(p).scanpdftext()
                    info = {'Name': name, 'Text': text}
                    uncompress_info.append(info)
    return uncompress_info


def init_index(index_name):
    """
    Initialize some files for storing the results retured by Function
    walk_file().
    :param index_name: A str. The name of current index setted by user.
    """

    # ToDo: 修改路径。用户创建索引名时就创建。
    # 轮换存三个扫描结果的文档，便于更新时对照。定期更新或手动更新索引，就覆写最早修改时间的
    directory = os.path.join("C:\\temp_index", str(index_name))
    os.makedirs(directory)
    num = 1
    for i in range(3):
        index = index_name + str(num) + ".txt"
        index_path = os.path.join(directory, index)
        with open(index_path, "w") as f:
            f.write("")
        num = num + 1


def walk_file(user_dirs, user_extensions, index_name):
    """
    Walk through the selected file_dirs and get the selected format
    files’info as a dictionary(Key:Value).
    :param user_dirs: A list. User's selected directories for scanning. e.g.
    ["C:/text", "D:/Programs"]
    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    :param index_name: A str. The name of current index setted by user.
    :return: A list, consists of dictionaries. For each dictionary, the
    Keys:'Name':file name; 'Extension': file format; 'Ctime': created time;
    'Mtime': last modified time; 'Path': absolute path of file; 'Size': the
    size of file(Unit:Mb); 'Text': extracted text(empty now).
    """

    contents = []
    for f_dir in user_dirs:  # if chose more than 1 file_dir
        for root, dirs, files in os.walk(f_dir):
            for f in files:
                extension = os.path.splitext(f)[1]
                if extension in user_extensions:
                    path = os.path.join(root, f)
                    name = os.path.splitext(f)[0]
                    ctime = timestamp2time(os.path.getctime(path))
                    mtime = timestamp2time(os.path.getmtime(path))
                    size = os.path.getsize(path) / float(1024 * 1024)  # Mb
                    info = {'Name': name, 'Extension': extension,
                            'Ctime': ctime, 'Mtime': mtime, 'Path': path,
                            'Size': size, 'Text': ''}
                    contents.append(info)
    # 轮换存三个扫描结果的文档，便于更新时对照。定期更新或手动更新索引，就覆写最早修改时间的
    pathname = os.path.join("C:\\temp_index", str(index_name))
    # 覆写最早修改的一个文档
    rewrite_early(index_dir=pathname, contents=contents)
    return contents


def get_text(contents, user_extensions):
    """
    Extract the text from selected directories and formats.
    :param contents: A list, consists of dictionaries. Acturally, It is the
    result returned by Function walk_file().
    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    :return: A list, consists of dictionaries. For each dictionary, the
    Keys:'Name':file name; 'Extension': file format; 'Ctime': created time;
    'Mtime': last modified time; 'Path': absolute path of file; 'Size': the
    size of file(Unit:Mb); 'Text': extracted text. Specially, the sign "'" in
    original text should be changed to sign "‘".
    """

    all_information = contents
    # ToDo: 文本过滤器
    scripts = ['.py', '.r', '.cpp']
    for i in contents:
        start_time = time.time()
        print(i['Path'])
        ext = os.path.splitext(i['Path'])[-1]
        # Extract text
        if ext == '.txt' or ext == '.md' or ext == '.yml':
            text = extracttext.TXTText(i['Path']).txttext()
            i['Text'] = text
        if ext in scripts:
            text = extracttext.TXTText(i['Path']).txttext()
            i['Text'] = text
        if ext == '.csv':
            text = extracttext.TXTText(i['Path']).csvtext()
            i['Text'] = text
            # extracttext.TXTText.rm_txt_files()
        if ext == '.xls' or ext == '.xlsx' or ext == '.xlsm':
            text = extracttext.TXTText(i['Path']).exceltext()
            i['Text'] = text
            # extracttext.TXTText.rm_txt_files()
        if ext == '.docx':
            text = extracttext.WordText(i['Path']).docxtext()
            i['Text'] = text
        if ext == '.doc' or ext == '.docm' or ext == '.rtf':
            text = extracttext.WordText(i['Path']).doctext()
            i['Text'] = text
        if ext == '.pptx':
            text = extracttext.PPTText(i['Path']).pptxtext()
            i['Text'] = text
        if ext == '.ppt' or ext == '.pptm':
            text = extracttext.PPTText(i['Path']).ppttext()
            i['Text'] = text
        if ext == '.odt' or ext == '.ods' or ext == '.odp':
            text = extracttext.ODFText(i['Path']).odftext()
            i['Text'] = text
        if ext == '.xml':
            text = extracttext.MarkupText(i['Path']).xmltext()
            i['Text'] = text
        if ext == '.html':
            text = extracttext.MarkupText(i['Path']).htmltext()
            i['Text'] = text
        if ext == '.tex':
            text = extracttext.MarkupText(i['Path']).textext()
            i['Text'] = text
        if ext == '.pdf':
            text = extracttext.PDFText(i['Path']).docpdftext()
            if text != '':
                text = text
                i['Text'] = text
            else:
                text = extracttext.PDFText(i['Path']).scanpdftext()
                i['Text'] = text
                # extracttext.PDFText.rm_spdf()
        if ext == '.tar' or ext == '.rar' or ext == '.zip' or ext == '.7z':
            uncompress.uncompress(file_path=i['Path'],
                                  user_extensions=user_extensions)
            uncompress.recursive(user_extensions=user_extensions)
            uncompress_info = get_uncompress_text()
            for ui in uncompress_info:
                u_name = str(ui['Name']) + ' (in) ' + i['Name']
                u_path = '(' + i['Path'] + ')'
                u_info = {'Name': u_name, 'Extension': i['Extension'],
                          'Ctime': i['Ctime'], 'Mtime': i['Mtime'],
                          'Path': u_path, 'Size': i['Size'],
                          'Text': ui['Text']}
                all_information.append(u_info)
            uncompress.rm_unzip_files()
        end_time = time.time()
        exec_time = end_time - start_time
        print(exec_time)
        print('\n')
    return all_information


def indexing(user_dirs, user_extensions, index_name):
    """
    The Main Function for indexing.
    :param user_dirs: A list. User's selected directories for scanning. e.g.
    ["C:/text", "D:/Programs"]
    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    :param index_name: A str. The name of current index setted by user.
    :return: A list, consists of dictionaries. For each dictionary, the
    Keys:'Name':file name; 'Extension': file format; 'Ctime': created time;
    'Mtime': last modified time; 'Path': absolute path of file; 'Size': the
    size of file(Unit:Mb); 'Text': extracted text. Specially, the sign "'" in
    original text should be changed to sign "‘".
    """

    directory = os.path.join("C:\\temp_index", index_name)
    if not os.path.exists(directory):
        init_index(index_name)
    contents = walk_file(user_dirs, user_extensions, index_name)
    index_information = get_text(contents=contents,
                                 user_extensions=user_extensions)
    return index_information


def renew_index(user_dirs, user_extensions, index_name):
    """
    Renew the index: renew the files' information; and based on the differences
    between new and old version files' information[returned result of Function
    walk_file()], remove the omitted ones, extract new texts, including added
    files or modified files.
    :param user_dirs: A list. User's selected directories for scanning. e.g.
    ["C:/text", "D:/Programs"]
    :param user_extensions: A list. User's selected formats for indexing. e.g.
    [".txt", ".xlsx"]
    :param index_name: A str. The name of current index setted by user.
    :return:removes: A list, consists of dictionaries. Its form is alike the
    returned result of Function walk_file()[empty text]. It contains the items
    that should be removed out from the current index.
            add_information: A list, consists of dictionaries. It is the
    returned result of Function get_text()[with extracted text]. It would be
    added to the current index.
    """

    directory = os.path.join("C:\\temp_index" + index_name)
    last_index = read_new(directory)
    newest_index = walk_file(user_dirs, user_extensions, index_name)
    # compare the differences
    adds = [x for x in newest_index if x not in last_index]
    removes = [y for y in last_index if y not in newest_index]
    add_information = get_text(contents=adds, user_extensions=user_extensions)
    return removes, add_information


if __name__ == "__main__":
    information = indexing(
        user_dirs=["D:\\othercode"],
        user_extensions=['.zip', '.xlsx'],
        index_name="test")
    print(information)
