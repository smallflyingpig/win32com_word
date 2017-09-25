# -*- coding: utf-8 -*-

# -------------------------
# lijiguo , 20170925, version 1.0
# -------------------------

import docx
import os
import re
import win32com.client as client

def normalize_MPEG_filename(filefolder, filename):
    #filefolder = u'E:\\lijiguo\\workspace\\python\\py_docx';
    #filename = u'test.doc';
    title_table_idx = 2;

    app = client.Dispatch('Word.Application')
    print("dispatch end")
    os.chdir(filefolder);
    print(os.getcwd());
    print(filename);
    file_path = os.path.join(os.getcwd(), filename);
    try:
        if not os.path.exists(file_path):
            return;
        doc=app.Documents.Open(os.path.join(os.getcwd(), filename))
        #app.Visible = True
        #app.ScreenUpdating = True

        if doc.Tables.Count<2:
            return;
        table = doc.Tables(title_table_idx);
        new_file_name = u'';
        title_str = str(table.Cell(1,1));
        idx = 1;#title_str.find(u'Title:');
        if idx>=0:
            new_file_name = str(table.Cell(1,2));
            doc.Close();
            new_file_name = new_file_name[:new_file_name.find('\r')];
            new_file_name = re.sub('[\/:*?"<>|\']', '-', new_file_name)
            new_file_name = re.sub(re.compile('\s+'), '', new_file_name)
            new_file_name.replace('\r', '');
            new_file_name.replace('\n', '');
            new_file_name = new_file_name + filename[filename.find('.'):];
            print(new_file_name);
            os.rename(filename, new_file_name);
        else:
            print(u'error: title is not matched');
            doc.Close();
    except:
        pass;
    
def normalize_folder(folder_path):
    # type: (object) -> object
    os.chdir(folder_path);
    for root,dirs,files in os.walk('./'):
        for file in files:
            name, ext = os.path.splitext(file);
            if name[0] == '.':
                continue;

            if ext == '.doc' or ext == '.docx':
                print(file+'...');
                normalize_MPEG_filename(folder_path, file);

        for dir_name in dirs:
            print("dir:"+dir_name);
            if len(dir_name)>=10:
                normalize_folder(os.path.join(folder_path, dir_name));


if __name__=='__main__':
    path_name = os.getcwd();
    normalize_folder(path_name);



