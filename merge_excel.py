import pandas as pd
from pathlib import Path
import logging
from itertools import chain
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.INFO)

class MergeExcel:
    """
    使用前请确保excel 有相同的列名称，否则会有数据上的错误
    """
    def __init__(self, path,sum=0):
        self.path = Path(path)
        self.sum=sum

    def files_path(self,exts=['xlsx','xls']):
        files=list()
        self.exts=exts
        for ext in exts:
                    # 获取当前目录下的所有xslx 文件,返回是个generate
            files.append(self.path.glob(f'*.{ext}'))
        files=list(chain(*files))
        return files

    def get_content(self):
        files=self.files_path()
        all_cont = pd.DataFrame()
        usecols= int(input('\033[1;36m please input the Column Number you want to use for this action:\033[0m').strip())
        usecols=range(usecols)
        for file in files:
            print('======='*10)
            logging.info(f'file name: {file}')
            # 读取每个excel 的各个sheetDefaults ： 第一页作为数据文件
            # 1 ：第二页作为数据文件
            # “Sheet1” ：第一页作为数据文件
            # [0,1，“SEET5”] ：第一、第二和第五作为作为数据文件
            # None ：所有表为作为数据文件
            content = pd.read_excel(file, sheet_name=None, encoding='utf8')
            # 获取所有的sheet name
            sheet_names = content.keys()
            #决定使用前几列

            if len(sheet_names) > 1:
                logging.info(f'this excel contains many sheets')
                for k, sheet_name in enumerate(sheet_names):
                    print(f'\033[1;33m No.{k} sheet name- {sheet_name} \033[0m')
                # 输入多个你想要合并的sheet name
                sheet_name_inputs = input('\033[1;36m please input the sheetnames you want to merge,split with ",":\033[0m').split(',')
                for sheet_name in sheet_name_inputs:
                    content = pd.read_excel(file, sheet_name=sheet_name.strip(), encoding='utf-8-sig',usecols=usecols)
                    logging.info(f'Sheet {sheet_name} own line number before: {content.shape[0]}')
                    content=content.dropna(how='all')
                    logging.info(f'Sheet {sheet_name} own columns: {content.columns}')
                    logging.info(f'Sheet {sheet_name} own line number after dropna: {content.shape[0]}')
                    self.sum += content.shape[0]
                    all_cont = pd.concat([all_cont, content], axis=0, sort=False)
            else:
                content = pd.read_excel(file, encoding='utf-8-sig',usecols=usecols)
                logging.info(f'this file own line number before: {content.shape[0]}')
                content=content.dropna(how='all')
                logging.info(f'this file own columns: {content.columns}')
                logging.info(f'this file own line number after dropna: {content.shape[0]}')
                self.sum += content.shape[0]
                all_cont = pd.concat([all_cont, content], axis=0, sort=False)
        logging.info(f'\033[1;36m All the sum of content is: {self.sum} \033[0m')
        return all_cont


if __name__ == '__main__':
    me = MergeExcel(r'D:\download_D\1230_小万歌曲标注\test')
    all_content = me.get_content()
    logging.info(f'\033[1;36m All content of merged shape is {all_content.shape}\033[0m')
    # output_name = Path(r'D:\download_D\1230_小万歌曲标注\test\res.csv')
    output_name = Path(r'D:\download_D\1230_小万歌曲标注\test\res.xlsx')
    if output_name.exists():
        output_name.unlink()
    # all_content.to_csv(output_name, index=None, header=True, mode='w', encoding='utf-8')
    all_content.to_excel(output_name, index=None, header=True, encoding='utf-8')
