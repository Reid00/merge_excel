import pandas as pd
from pathlib import Path


class MergeExcel:
    def __init__(self, path):
        self.path = Path(path)

    def get_files(self, ext='xlsx'):
        self.ext = ext
        # 获取当前目录下的所有xslx 文件,返回是个generate
        files = self.path.glob(f'*.{self.ext}')
        all_cont = pd.DataFrame()
        sum = 0
        for file in files:
            print(f'file name {file}')
            # 读取每个excel 的各个sheet
            content = pd.read_excel(file, sheet_name=None, encoding='utf8')
            # 获取所有的sheet name
            sheet_names = content.keys()
            if len(sheet_names) > 1:
                print(f'this excel contains many sheets')
                for k, sheet_name in enumerate(sheet_names):
                    print(f'{k}-{sheet_name}')
                sheet_name_input = input('please input the sheetname you want to merge:')
                content = pd.read_excel(file, sheet_name=sheet_name_input, encoding='utf8')
                print(f'file line number {content.shape[0]}')
                sum += content.shape[0]
                all_cont = pd.concat([all_cont, content], axis=0, sort=False)
            else:
                content = pd.read_excel(file, encoding='utf8')
                print(f'file line number {(content.shape[0])}')
                sum += content.shape[0]
                all_cont = pd.concat([all_cont, content], axis=0, sort=False)
        print(f'all the sum of content is {sum}')
        return all_cont


if __name__ == '__main__':
    me = MergeExcel(r'C:\Users\v-baoz\Downloads\1029_poems')
    all_content = me.get_files()
    print(f'all content of merged shape is {all_content.shape}')
    output_name = Path(r'C:\Users\v-baoz\Downloads\1029_poems\res.csv')
    if output_name.exists():
        output_name.unlink()
    all_content.to_csv(output_name, index=None, header=True, mode='w')
