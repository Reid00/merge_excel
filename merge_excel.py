import numpy as np
import pandas as pd
from pathlib import Path
import logging
from itertools import chain
import re
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.INFO,filemode='w',
# filename='log.log'
)
logger=logging.getLogger(__name__)

class MergeExcel:
    """
    使用前请确保excel 有相同的列名称，否则会有数据上的错误
    """
    def __init__(self, path,sum=0,exts=['xlsx','xls']):
        self.path = Path(path)
        self.sum=sum
        self.exts=exts

    def files_path(self):
        """
        获取需要处理的文件路径，去除到处的结果res
        """
        files = list()
        for ext in self.exts:
            # 获取当前目录下的所有xslx 文件,返回是个generate
            files.append(self.path.glob(f'*.{ext}'))
        files = list(chain(*files))
        files = [file for file in files if file.stem != 'res']
        return files

    def get_content(self):
        """
        获取每个excel 里面的内容
        """
        files=self.files_path()
        all_cont = pd.DataFrame()
        usecols= int(input('\033[1;36m please input the Column Number you want to use for this action:\033[0m').strip())
        usecols=range(usecols)
        for file in files:
            print('======='*10)
            logger.info(f'file name: {file}')
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
                logger.info(f'this excel contains many sheets')
                for k, sheet_name in enumerate(sheet_names):
                    print(f'\033[1;33m No.{k} sheet name- {sheet_name} \033[0m')
                # 输入多个你想要合并的sheet name
                sheet_name_inputs = input('\033[1;36m please input the sheetnames you want to merge,split with ",":\033[0m').split(',')
                for sheet_name in sheet_name_inputs:
                    if sheet_name == '':
                        logger.info(f'nothing selected')
                        continue
                    content = pd.read_excel(file, sheet_name=sheet_name.strip(), encoding='utf-8-sig',usecols=usecols)
                    logger.info(f'Sheet {sheet_name} own line number before: {content.shape[0]}')
                    content=content.dropna(how='all')
                    logger.info(f'Sheet {sheet_name} own columns: {content.columns}')
                    logger.info(f'Sheet {sheet_name} own line number after dropna: {content.shape[0]}')
                    self.sum += content.shape[0]
                    all_cont = pd.concat([all_cont, content], axis=0, sort=False)
            else:
                content = pd.read_excel(file, encoding='utf-8-sig',usecols=usecols)
                logger.info(f'this file own line number before: {content.shape[0]}')
                content=content.dropna(how='all')
                logger.info(f'this file own columns: {content.columns}')
                logger.info(f'this file own line number after dropna: {content.shape[0]}')
                self.sum += content.shape[0]
                all_cont = pd.concat([all_cont, content], axis=0, sort=False)
        logger.info(f'\033[1;36m All the sum of content is: {self.sum} \033[0m')
        return all_cont

    def value_counts_info(self):
        """
        value_counts 的一些方法使用，可以用pivot 代替
        """
        dataframe = pd.DataFrame({
            'name': ['Reid', 'Reid', 'Knight', 'Tom'],
            'times': [1, 2, 2, 3]
        })
        # 统计name出现次数
        counts=dataframe['name'].value_counts()
        print(counts)
        # 统计name出现两次以上的
        cnt_gt2= counts[counts>1]
        print(cnt_gt2)
        # 统计name 包含i 的有哪些
        name_contains_i= dataframe['name'].str.contains('i')
        print(name_contains_i)

    def sort_according_lst(self):
        """
        根据一个list 顺序对dataframe 的某一列进行排序
        
        """
        order_lst=['b','a','c']
        data=pd.DataFrame({
            'words':list('abc'),
            'number':list('123')
        })
        print(data)
        #相等的情况下，可以使用 reorder_categories和 set_categories方法；
        # inplace = True，使 set_categories生效
        data['words']=data['words'].astype('category')
        data['words'].cat.set_categories(order_lst,inplace=True)
        data.sort_values('words',inplace=True)
        print(f'list 数目相同: \n {data}')
        print('========='*10)
        # list的元素比较多的情况下， 可以使用set_categories方法；
        # list的元素比较少的情况下， 也可以使用set_categories方法，但list中没有的元素会在DataFrame中以NaN表示。
        order_lst=['b','c']
        data['words'].cat.set_categories(order_lst,inplace=True)
        data.sort_values(by=['words'],inplace=True)
        print(f'list 数目比较少: \n{data}')
        print('========='*10)

        data=pd.DataFrame({
            'words':list('abc'),
            'number':list('123')
        })
        order_lst=['b','c','a','d']
        data['words']=data['words'].astype('category')
        data['words'].cat.set_categories(order_lst,inplace=True)
        data.sort_values(by=['words'],inplace=True)
        print(f'list 数目比较多: \n{data}')
        print('========='*20)

    def rm_blank(self,dataframe,*columns):
        """
        去除指定列中所有的空格
        """
        for col in columns:
            # breakpoint()
            dataframe[col]=dataframe[col].astype(str).apply(lambda x: re.sub(r'\s+','',x))
            # dataframe[col]=dataframe[col].apply(lambda x: re.sub(r'\s+','',str(x)))
            # dataframe[col]=dataframe[col].astype(str).str.replace(r'\s','',regex=True)
            dataframe.replace('nan','',inplace=True)
        return dataframe
    def rm_strip(self,dataframe,*columns):
        """
        去除指定列中前后的空格
        """
        for col in columns:
            dataframe[col]=dataframe[col].astype(str).apply(lambda x:x.strip())
            dataframe.replace('nan','',inplace=True)
        return dataframe

if __name__ == '__main__':
    me = MergeExcel(r'D:\download_D\1230_小万内容清洗\0109')
    all_content = me.get_content()
    columns_rm_blank=['专辑名称','歌曲名称','上下']
    columns_rm_strip=['专辑Tag','歌曲tag']
    all_content=me.rm_blank(all_content,*columns_rm_blank)
    all_content=me.rm_strip(all_content,*columns_rm_strip)
    all_content.drop_duplicates(subset=['歌曲id'],keep='first',inplace=True)
    logger.info(f'\033[1;36m All content of merged shape is {all_content.shape}\033[0m')
    output_name = Path(r'D:\download_D\1230_小万内容清洗\0109\res.xlsx')
    if output_name.exists():
        output_name.unlink()
    # all_content.to_csv(output_name, index=None, header=True, mode='w', encoding='utf-8')
    all_content.to_excel(output_name, index=None, header=True, encoding='utf-8')
    print(r'jobs done')
