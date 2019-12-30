import pandas as pd
from pathlib import Path
import logging
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.INFO)

class MergeExcel:
    """
    使用前请确保excel 有相同的列明，否则会有数据上的错误
    """
    def __init__(self, path):
        self.path = Path(path)

    def get_files(self, ext='xlsx'):
        self.ext = ext
        # 获取当前目录下的所有xslx 文件,返回是个generate
        files = self.path.glob(f'*.{self.ext}')
        all_cont = pd.DataFrame()
        sum = 0
        for file in files:
            logging.info(f'file name {file}')
            # 读取每个excel 的各个sheet
            content = pd.read_excel(file, sheet_name=None, encoding='utf8')
            # 获取所有的sheet name
            sheet_names = content.keys()
            #决定使用前几列
            usecols= int(input('\033[1;34m please input the Number you want to use of all excel: \033[0m').strip())
            usecols=range(usecols)
            if len(sheet_names) > 1:
                logging.info(f'this excel contains many sheets')
                for k, sheet_name in enumerate(sheet_names):
                    print(f'\033[1;33m No.{k} sheet name- {sheet_name} \033[0m')
                # 输入多个你想要合并的sheet name
                sheet_name_inputs = input('\033[1;34m please input the sheetnames you want to merge,split with ",":\033[0m').split(',')
                for sheet_name in sheet_name_inputs:
                    content = pd.read_excel(file, sheet_name=sheet_name.strip(), encoding='utf-8-sig',usecols=usecols)
                    logging.info(f'the columns of this sheet: {content.columns}')
                    logging.info(f'file line number: {content.shape[0]}')
                    logging.info(f'file shape is:  {content.shape}')
                    sum += content.shape[0]
                    all_cont = pd.concat([all_cont, content], axis=0, sort=False)
            else:
                content = pd.read_excel(file, encoding='utf8')
                logging.info(f'the columns of this sheet: {content.columns}')
                logging.info(f'file line number: {(content.shape[0])}')
                logging.info(f'file shape is:  {content.shape}')
                sum += content.shape[0]
                all_cont = pd.concat([all_cont, content], axis=0, sort=False)
        logging.info(f'\033[1;34m All the sum of content is: {sum} \033[0m')
        return all_cont


if __name__ == '__main__':
    me = MergeExcel(r'D:\download_D\1230_小万歌曲标注\test_many_sheet')
    all_content = me.get_files()
    logging.info(f'\033[1;34m All content of merged shape is {all_content.shape} \033[0m')
    output_name = Path(r'D:\download_D\1230_小万歌曲标注\test_many_sheet\res.csv')
    if output_name.exists():
        output_name.unlink()
    all_content.to_csv(output_name, index=None, header=True, mode='w', encoding='utf-8-sig')
