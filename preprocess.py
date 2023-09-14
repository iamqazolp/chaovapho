import os
import re
import time
import sys
import shutil
import subprocess
import pandas as pd
import glob
import textract
import docx

def process_archive(inputFolder, passwords):
    archive_exts = ['zip', 'tar', 'tar.gz', 'tar.bz2', 'tgz', 'rar', '7z']
    bin7z = os.path.join(os.getcwd(), "bin", "7z.exe")
    uncompressedFolder = os.path.join(inputFolder, "uncompressed")
    for ext in archive_exts:
        archive_files = glob.glob(os.path.join(inputFolder, "**", "*."+ext), recursive=True)
        for f in archive_files:
            unzip_ok = False
            for password in passwords:
                outPath = os.path.join(uncompressedFolder, os.path.basename(f) + ".tmp")
                cmd = f'{bin7z} e {f} -o{outPath} -r -y -p{password}'
                p = subprocess.run(cmd, capture_output=True, text=True)
                if p.stderr == None or p.stderr == '':
                    print('Unziped ', f)
                    unzip_ok = True
                    break
            if not unzip_ok:
                shutil.rmtree(outPath, ignore_errors=True)
            if 'Wrong password' in p.stderr:
                wrong_pass_path = os.path.join(inputFolder, "wrong_password")
                os.makedirs(wrong_pass_path, exist_ok=True)
                shutil.copy(f, wrong_pass_path)
            

    while True:
        archive_files = []
        for ext in archive_exts:
            archive_files += glob.glob(os.path.join(inputFolder,"uncompressed", "**", "*."+ext), recursive=True)
        print(archive_files)
        if (len(archive_files) == 0):
            break
        for f in archive_files:
            unzip_ok = False
            for password in passwords:
                cmd = f'{bin7z} e {f} -o{f}.tmp -y -r -p{password}'
                p = subprocess.run(cmd, capture_output=True, text=True)
                if p.stderr == None or p.stderr == '':
                    print('Unziped ', f)
                    unzip_ok = True
                    break
            if not unzip_ok:
                shutil.rmtree(outPath)
            if 'Wrong password' in p.stderr:
                wrong_pass_path = os.path.join(inputFolder, "wrong_password")
                os.makedirs(wrong_pass_path)
                shutil.copy(f, wrong_pass_path)
            os.remove(f)

def process_msoffice(inputFolder):
    word_exts = ['doc', 'docx']
    excel_exts = ['xls', 'xlsx']
    print('Processing Word')
    for ext in word_exts:
        ms_files = glob.glob(os.path.join(inputFolder, "**", "*."+ext), recursive=True)
        for f in ms_files:
            with open(f + ".txt", "wb") as file:
                file.write(textract.process(r'{}'.format(f)))
    # word.Quit()

    print('Processing Excel')
    for ext in excel_exts:
        ms_files = glob.glob(os.path.join(inputFolder, "**", "*."+ext), recursive=True)
        for f in ms_files:
           print('Processing ' + f)
           df = pd.read_excel(f, sheet_name=None)
           i=0
           for key in df:
               i += 1
               df[key].to_csv(f + "." + str(i) + ".csv", index=False)

def find_keywords(keywords, inputFolder, outputFolder):
    bin_grep = os.path.join(os.getcwd(), "bin", "grep.exe")
    cmd = f'{bin_grep} -arilE "{keywords}" {inputFolder}'
    print('Searching for keywords: ', keywords)
    p = subprocess.run(cmd, capture_output=True, text=True)
    out = os.path.join(outputFolder, keywords.replace("|", "_"))
    os.makedirs(out, exist_ok=True)
    results = [re.sub(r'\.xls\..*\.csv', '.xls', re.sub(r'\.xlsx\..*\.csv', '.xlsx', line.replace(".doc.txt", ".doc").replace(".docx.txt", ".docx"))) for line in p.stdout.splitlines()]
    with open(out + ".txt", "w") as f:
        f.writelines(results)
    for line in results:
        shutil.copy(line, out)
    return results
        
if __name__ == '__main__':
    df = pd.read_excel('input.xlsx')
    inputFolder = os.path.abspath(df['ThuMucDauVao'][0])
    
    outputFolder = os.path.abspath(df['ThuMucKetQua'][0])

    passwords = [x for x in df['MatKhauGiaiNen'].dropna()]
    keywords = [x for x in df['TuKhoaTimKiem'].dropna()]
    print('inputFolder: ', inputFolder)
    print('outputFolder: ', outputFolder)
    print('passwords: ', passwords)
    print('keywords: ', keywords)
    process_archive(inputFolder, passwords)
    process_msoffice(inputFolder)
    for keys in keywords:
        find_keywords(keys, inputFolder, outputFolder)