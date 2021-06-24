

import shutil, os
from os import walk

original_patch = r'C:/Git/Display_ppt/Drop/'
target_path = r'C:/Git/Display_ppt/copy/'


def copy_file(source, target):
    original = []  # get the list of file in directory 'path'
    for (dirpath, dirnames, original_filenames) in walk(source):
        original.extend(original_filenames)
        break
    print('file list:', original)

    copy = []  # get the list of file in directory 'path'
    for (dirpath, dirnames, copy_filenames) in walk(target):
        copy.extend(copy_filenames)
        break
    print('file list:', copy)

    for name in original:   #if file not presnt in destination folder copy it
        present=False
        for dest_name in copy:
            if name == dest_name:
                present=True
        if not present:
            try:
                shutil.copyfile(source + name, target + name)
                print("file copied", source + name, " to ", target + name )
            except:
                print('failed to copy')
        else:
            print("file already present", source + name )

    for dest_name in copy: #if file not present in source directory erase it in destination directory
        present = False
        for name in original:
            if name == dest_name:
                present = True
        if not present:
            try:
                os.remove(target + dest_name)
                print("file deleted from", target + dest_name)
            except FileNotFoundError:
                print('file not found, failed to delete')

        else:
            print("file present", source + name)






copy_file(original_patch,target_path)