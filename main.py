#!/bin/python
# _*_coding:utf-8_*_
# oc.py
import os
import re
import subprocess
import zipfile
from datetime import datetime
from typing import List, Dict
from tqdm import tqdm
from shlex import quote

import toml

FILES_PATH = ['./']
COMPR_LEVL = 4  # 1 - 10

allow_ext = ['docx', 'xlsx', 'pptx']
compress_level = {
    'tiff': ['-c', 0, 1, 2, 3, 4, 5, 6, 7, 8],
    'jpg': ['-q', 1, 100],
    'png': ['-clevel', 1, 10]
}
max_level = 10
min_level = 0
def_args = ['-no_auto_ext', '-overwrite']
config = {}


def find_nconvert(ncver_path: str):
    if not os.path.isdir(ncver_path):
        with os.popen(' '.join([ncver_path, '-help'])) as _f:
            while _f.readable():
                if 'nconvert' in _f.readline().lower():
                    return ncver_path
        return None

    for current, dirs, files in os.walk(ncver_path):
        for filename in files:
            if filename != 'nconvert':
                continue
            ncver = os.path.join(current, filename)
            with os.popen(' '.join([ncver, '-help'])) as _f:
                while _f.readable():
                    if 'nconvert' in _f.readline().lower():
                        return ncver
    return None


def load_config():
    return toml.load('config.toml')


def save_config(config: Dict):
    with open('config.toml', 'w') as _f:
        toml.dump(config, _f)


def init():
    '''check nconvert installed'''
    global config
    try:
        config = load_config()
    except Exception:
        print('bad configuration file. use default configuration.')
    if 'nconvert_path' not in config:
        config['nconvert_path'] = './'
    ncver_path = find_nconvert(config['nconvert_path'])
    if ncver_path:
        config['nconvert_path'] = ncver_path
    else:
        print(
            '''can't find "nconvert", please visit https://www.xnview.com/en/nconvert/ to download and put it in the program directory''')
        return False

    if 'cache_path' not in config:
        config['cache_path'] = './cache/'
    if not os.path.isdir(config['cache_path']):
        os.makedirs(config['cache_path'])
    save_config(config)
    return True


def get_file_ext(filename: str):
    r_name = re.search('(.+?)(\.[^.]+?)$', filename)
    if not r_name:
        return r_name
    return r_name.group(1), r_name.group(2)


def gen_comlvl_bmp(level: int = 10) -> (List[str]):
    return ['-c', str(round(1 / max_level * level))]


def gen_comlvl_gif(level: int = 10) -> (List[str]):
    return ['-c', '0']


def gen_comlvl_tiff(level: int = 10) -> (List[str]):
    return ['-c', str(round(8 / max_level * level))]


def gen_comlvl_jpg_fpx_wic(level: int = 10) -> (List[str]):
    return ['-q', str(round(100 / max_level * (max_level - level))), '-opthuff', '-rmeta']


def gen_comlvl_png(level: int = 10) -> (List[str]):
    return ['-clevel', str(round(9 / max_level * level))]


def gen_comlvl(ext, level: int = 6) -> (List[str], None):
    if level > max_level or level < min_level:
        raise ValueError('compression level must be > 0 and <= 10')
    if ext == 'bmp':
        return gen_comlvl_bmp(level)
    elif ext == 'gif':
        return gen_comlvl_gif(level)
    elif ext == 'tif' or ext == 'tiff':
        return gen_comlvl_tiff(level)
    elif ext == 'jpg' or ext == 'jpeg' or ext == 'fpx' or ext == 'wic':
        return gen_comlvl_jpg_fpx_wic(level)
    elif ext == 'png':
        return gen_comlvl_png(level)
    else:
        return None


def images_compress(dirname, lvl: int = 6):
    for current, dirs, files in os.walk(dirname):
        for filename in tqdm(files):
            image_compress(current, filename, lvl)


def image_compress(path: str, filename: str, lvl: int = 6):
    '''
    image compress
    '''
    global config
    args_new = def_args.copy()
    imgfile = get_file_ext(filename)
    args = gen_comlvl(imgfile[1][1:], lvl)
    if not args:
        print(
            'the file "' + filename + '" format does not support compression at the moment, feel free to submit issus.')
        return
    args_new.extend(args)
    args_new.append(os.path.join(path, filename))
    args_final = [config['nconvert_path']]
    for i in args_new:
        args_final.append(quote(i))
    result = subprocess.run(
        args_final,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL)


def unzip(filename, outname: str = None):
    '''
    unzip zip file
    '''
    f_zip = zipfile.ZipFile(filename)
    if not outname:
        outname = 'unzip_' + filename + '_' + datetime.now().strftime('%Y%m%d')
    if os.path.isdir(outname):
        print('the temporary file ' + filename + ' already exists, skip')
        return False
    else:
        os.mkdir(outname)
    # if outname.endswith('/'):
    #     outname += '/'
    for names in f_zip.namelist():
        f_zip.extract(names, outname)
    f_zip.close()
    return True


def rezip(srcname, filename: str = None):
    '''
    unzip zip file
    '''
    if not filename:
        filename = 'zipfile' + datetime.now().strftime('%Y%m%d') + '.zip'
    f_zip = zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED)
    for current, dirs, files in os.walk(srcname):
        for filename in files:
            f_zip.write(os.path.join(current, filename), os.path.join(re.sub(srcname, '', current), filename))
    f_zip.close()


def run():
    global config
    files_list = []
    for filepath in FILES_PATH:
        for current, dirs, files in os.walk(filepath):
            for filename in files:
                ext_file = get_file_ext(filename)
                if not ext_file or ext_file[1][1:] not in allow_ext:
                    continue
                files_list.append([current, filename])
    for path, filename in tqdm(files_list):
        # path, filename = fileinfo[0], fileinfo[1]
        outname = os.path.join(config['cache_path'], 'unzip_' + filename + '_' + datetime.now().strftime('%Y%m%d'))
        unzip(os.path.join(path, filename), outname)
        images_compress(os.path.join(outname, '/word/media/'))
        rezip(outname, os.path.join(path, 'compressed_' + filename))
    # get file ext
    # if not r_filename:
    #     print('the ' + filename + ' file format is not supported, skip')
    # print(r_filename.group())


# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

if __name__ == '__main__':
    if not init():
        exit(1)
    run()
