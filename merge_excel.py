import pandas as pd
from pathlib import Path


class MergeExcel:
    def __init__(self, path):
        self.path = Path(path)

    def get_files(self, ext='xlsx'):
        self.ext = ext
        # 获取当前目录下的所有xslx 文件,返回是个generate
        files = self.path.glob(f'*.{self.ext}')
        for file in files:
            print(file)


if __name__ == '__main__':
    me = MergeExcel(r'C:\Users\v-baoz\Downloads\1024_poems')
    me.get_files()
