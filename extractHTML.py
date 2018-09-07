# -*- coding:utf-8 -*-

__author__ = "yooongchun"
__email__ = "yooongchun@foxmail.com"
__wechat__ = '18217235290'

'''
This program is designed for parsing webpages and extracting information.
Just for some specific webpages.

'''
import requests
from bs4 import BeautifulSoup as bs
import re
import xlrd
from xlutils.copy import copy
import os
import xlwt
import threading
import multiprocessing


# parser
class HTML_Parser(object):
    def __init__(self, url, excel_path):
        self.__url = url
        self.__excel_path = excel_path

    def __query_html(self):
        response = None
        for i in range(3):
            try:
                print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'query url:', self.__url)
                response = requests.get(url=self.__url, timeout=15)
                if response.status_code == 200:
                    break
            except Exception:
                pass
        if response is None:
            return
        try:
            html = response.content.decode(response.encoding)
            return html
        except Exception:
            # print('decoding error.')
            return

    def __parseHTML(self):
        html = self.__query_html()
        if html is None:
            print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'html is None,return.')
            # error log
            with open('error.log', 'a', encoding='utf-8')as fp:
                fp.write('html is None:' + self.__url + '\n')

            return None
        soup = bs(html, 'lxml')
        tds = soup.find_all('td') + soup.find_all('TD')
        # print('searching legend symbol...')
        # legend
        target_letter = None
        for td in tds:
            text = td.text
            te = ''
            for t in text.split('\n'):
                te += t
            te = re.sub(r'\s+', ' ', te)
            if re.search(r'\s*\([a-z]\)\s*Security or a portion of the security is on loan at period end', te):
                target_letter = re.findall(r'^\s*\([a-z]\)', te)[0]
                break
        if target_letter is None:
            for index, td in enumerate(tds):
                if index > len(tds) - 4:
                    break
                text = td.text + tds[index + 1].text + tds[index + 2].text + tds[index + 3].text
                te = ''
                for t in text.split('\n'):
                    te += t
                te = re.sub(r'\s+', ' ', te)
                if re.search(r'\s*\([a-z]\)\s*Security or a portion of', te):
                    target_letter = re.findall(r'\([a-z]\)', te)[0]
                    break
        if target_letter is None:
            ps = soup.find_all('p') + soup.find_all('P')
            for index, p in enumerate(ps):
                text = p.text
                te = re.sub(r'\s+|\n+', ' ', text)
                if re.search(r'\s*\([a-z]\)\s*Security or a portion of the security is on loan at period end', te):
                    target_letter = re.findall(r'^\s*\([a-z]\)', te)[0]
                    break
            if target_letter is None:
                for index, p in enumerate(ps):
                    if index > len(ps) - 4:
                        break
                    text = p.text + ps[index + 1].text + ps[index + 2].text + ps[index + 3].text
                    te = re.sub(r'\s+|\n+', ' ', text)
                    if re.search(r'\s*\([a-z]\)\s*Security or a portion of', te):
                        target_letter = re.findall(r'^\s*\([a-z]\)', te)[0]
                        break
        if not target_letter:
            print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'legend symbol not find:',
                  self.__url)
            # error log
            with open('error.log', 'a', encoding='utf-8')as fp:
                fp.write('legend symbol not find:' + self.__url + '\n')
            return None
        print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'find legend symbol:',
              target_letter)

        # group target td
        target_tds = []
        flag = False
        for index, td in enumerate(tds):
            text = td.text
            text = re.sub(r'\s+|\n+', ' ', text)
            if re.search(r'continued', text) \
                    or re.search('\s+by\s+|By\s+|\s+by\n', text) \
                    or re.search(r'\s+of\s+', text) \
                    or re.search(r'\s+a\s+', text) \
                    or re.search(r'\s+is\s+', text) \
                    or re.search(r'\s+for\s+', text) \
                    or re.search(r'\s+as\s+', text) \
                    or re.search(r'\s+also\s+', text) \
                    or re.search(r'\s+the\s+', text) \
                    or re.search(r'\s+and\s+', text) \
                    or re.search(r'\s+to\s+', text) \
                    or re.search(r'\s+may\s+', text) \
                    or re.search(r'value', text.lower()) \
                    or re.search(r'share', text.lower()):
                continue
            if re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', text):
                target_tds.append(text)
                flag = True
                continue
            if target_letter in text and not re.search(r'^\s*\([a-z]\)', text):
                target_tds.append(text)
                continue
            if flag and re.search(r'^\s*[a-zA-Z][a-zA-Z\s.\-&,]+', text):
                target_tds.append(text)
        # fund name
        soup = bs(html, 'lxml')
        ps = soup.find_all('p') + soup.find_all('P')
        fund_name = None

        # target
        Type1 = None
        Type2 = None
        Type3 = None
        target = []
        for index, ttd in enumerate(target_tds):
            if index < len(target_tds) - 3 and \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', ttd) and \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 1]) and \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 2]) and not \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 3]):
                Type1 = re.findall(r'[a-zA-Z\s]+', ttd)[0]
                Type1 = re.sub(r'^\s+|\s+$', '', Type1)
                # fund name
                for ind, p in enumerate(ps):
                    if ind > len(ps) - 3:
                        break
                    text = p.text + ps[ind + 1].text + ps[ind + 2].text
                    text = re.sub(r'\s+|\n+', ' ', text)
                    if re.search(r'Fidelity[a-zA-Z\s®]+Fund', text):
                        fund_name = re.findall(r'Fidelity[a-zA-Z\s®]+Fund', text)[0]
                        ps = ps[ind + 1:]
                        break
                if fund_name is None:
                    fund_name = 'not-find'
                continue
            if index < len(target_tds) - 2 and \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', ttd) and \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 1]) and not \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 2]):
                Type2 = re.findall(r'[a-zA-Z\s]+', ttd)[0]
                Type2 = re.sub(r'^\s+|\s+$', '', Type2)
                continue
            if index < len(target_tds) - 1 and \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', ttd) and not \
                    re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 1]):
                Type3 = re.findall(r'[a-zA-Z\s]+', ttd)[0]
                Type3 = re.sub(r'^\s+|\s+$', '', Type3)
                continue
            if Type1 and Type2 and Type3:
                cnt = index + 1
                while cnt < len(target_tds) and re.search(r'^\s*Class', target_tds[cnt]):
                    if target_letter in target_tds[cnt]:
                        ttd += target_tds[cnt]
                        break
                    cnt += 1
                ttd = re.sub(r'\s+|\n+', ' ', ttd)
                if target_letter in ttd and not re.search(r'^\s*\([a-z]\)', ttd):
                    if re.search(r'\d%|\d/', ttd):
                        continue
                    if re.search(r'^\s*Class', ttd):
                        continue
                    target.append({'fund': fund_name, 'type1': Type1, 'type2': Type2, 'type3': Type3, 'name': ttd})
        # Just Type1 & Type3
        if len(target) < 1:
            Type1 = None
            Type3 = None
            for index, ttd in enumerate(target_tds):
                if index < len(target_tds) - 2 and \
                        re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', ttd) and \
                        re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 1]) and not \
                        re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 2]):
                    Type1 = re.findall(r'[a-zA-Z\s]+', ttd)[0]
                    Type1 = re.sub(r'^\s+|\s+$', '', Type1)
                    # fund name
                    for ind, p in enumerate(ps):
                        if ind > len(ps) - 3:
                            break
                        text = p.text + ps[ind + 1].text + ps[ind + 2].text
                        text = re.sub(r'\s+|\n+', ' ', text)
                        if re.search(r'Fidelity[a-zA-Z\s®]+Fund', text):
                            fund_name = re.findall(r'Fidelity[a-zA-Z\s®]+Fund', text)[0]
                            ps = ps[ind + 1:]
                            break
                    if fund_name is None:
                        fund_name = 'not-find'
                    continue
                if index < len(target_tds) - 1 and \
                        re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', ttd) and not \
                        re.search(r'^\s*[a-zA-Z][a-zA-Z\s]+.{1,2}\s*\d+\.\d+%\s*$', target_tds[index + 1]):
                    Type3 = re.findall(r'[a-zA-Z\s]+', ttd)[0]
                    Type3 = re.sub(r'^\s+|\s+$', '', Type3)
                    continue
                if Type1 and Type3:
                    cnt = index + 1
                    while cnt < len(target_tds) and re.search(r'^\s+Class', target_tds[cnt]):
                        if target_letter in target_tds[cnt]:
                            ttd += target_tds[cnt]
                            break
                        cnt += 1
                    ttd = re.sub(r'\s+|\n+', ' ', ttd)
                    if target_letter in ttd and not re.search(r'^\s*\([a-z]\)', ttd):
                        if re.search(r'\d%|\d/', ttd):
                            continue
                        if re.search(r'^\s*Class', ttd):
                            continue
                        target.append({'fund': fund_name, 'type1': Type1, 'type2': '', 'type3': Type3, 'name': ttd})
        return target

    def __save2Excel(self):
        try:
            INFO = self.__parseHTML()
        except Exception as e:
            print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'parse HTML error:', e)
            # error log
            with open('error.log', 'a', encoding='utf-8')as fp:
                fp.write('parse HTML error:' + self.__url + '\n')
            return
        if INFO is None or len(INFO) < 1:
            print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'parser finds nothing:',
                  self.__url)
            # error log
            with open('error.log', 'a', encoding='utf-8')as fp:
                fp.write('parser finds nothing:' + self.__url + '\n')
            return
        # Save to file
        print('process/thread:', os.getpid(), '-', threading.current_thread().name, 'save to file:', self.__excel_path)
        book = xlrd.open_workbook(self.__excel_path)
        sheet = book.sheet_by_name('Sheet1')
        rows = sheet.nrows
        copy_book = copy(book)

        sheet_copy = copy_book.get_sheet('Sheet1')
        for index, info in enumerate(INFO):
            value = sheet.cell_value(1, 0)
            sheet_copy.write(index + rows - 1, 0, value)
            value = sheet.cell_value(1, 1)
            sheet_copy.write(index + rows - 1, 1, value)
            value = sheet.cell_value(1, 2)
            sheet_copy.write(index + rows - 1, 2, value)
            value = info['fund']
            sheet_copy.write(index + rows - 1, 3, value)
            value = info['type1']
            sheet_copy.write(index + rows - 1, 4, value)
            value = info['type2']
            sheet_copy.write(index + rows - 1, 5, value)
            value = info['type3']
            sheet_copy.write(index + rows - 1, 6, value)
            value = info['name']
            sheet_copy.write(index + rows - 1, 7, value)
        copy_book.save(self.__excel_path)

    def run(self):
        self.__save2Excel()


# load demo Excel
def load_Excel(path, sheet_name):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name(sheet_name)
    rows = sheet.nrows
    info = []
    for row in range(rows):
        if row == 0:
            continue
        fund_name = sheet.cell_value(row, 0)
        report_date = sheet.cell_value(row, 1)
        report_date = xlrd.xldate_as_tuple(report_date, 0)
        report_date = str(report_date[0]) + '/' + str(report_date[1]) + '/' + str(report_date[2])
        filing_date = sheet.cell_value(row, 2)
        filing_date = xlrd.xldate_as_tuple(filing_date, 0)
        filing_date = str(filing_date[0]) + '/' + str(filing_date[1]) + '/' + str(filing_date[2])
        url = sheet.cell_value(row, 3)
        info.append(
            {'fund_name': str(fund_name),
             'report_date': str(report_date),
             'filing_date': str(filing_date),
             'url': url})
    return info


# add Excels
def add_Excel(info):
    if not os.path.isdir('Excels'):
        os.mkdir('Excels')
    for one in info:
        date = one['report_date'].split('/')
        base_name = one['fund_name'] + '_' + date[0] + date[1] + '.xls'
        full_path = os.path.join('Excels', base_name)
        book = xlwt.Workbook()
        sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
        # Header
        sheet.write(0, 0, 'Fund_series')
        sheet.write(0, 1, 'Period of Report')
        sheet.write(0, 2, 'Filing Date')
        sheet.write(0, 3, 'Fund-Name')
        sheet.write(0, 4, 'Type1')
        sheet.write(0, 5, 'Type2')
        sheet.write(0, 6, 'Type3')
        sheet.write(0, 7, 'Stock')
        # Common info
        sheet.write(1, 0, one['fund_name'])
        sheet.write(1, 1, one['report_date'])
        sheet.write(1, 2, one['filing_date'])

        book.save(full_path)


# multiple threads
def batch_processor(info):
    thread_pool = []
    for index, one in enumerate(info):
        url = one['url']
        date = re.split(r'/', one['report_date'])
        base_name = one['fund_name'] + '_' + date[0] + date[1] + '.xls'
        full_path = os.path.join('Excels', base_name)
        parser = HTML_Parser(url, full_path)
        th = threading.Thread(target=parser.run)
        # print('running thread:', th.name)
        th.start()
        thread_pool.append(th)
    for th in thread_pool:
        # print('waiting for thread:', th.name)
        th.join()


# multiple processors
def multi_processor_run(func, info):
    pool = multiprocessing.Pool(processes=4)
    cnt = 0
    if os.path.isfile('error.log'):
        os.remove('error.log')
    while cnt < len(info):
        rear = cnt + 20
        if rear > len(info):
            rear = len(info)
        batch = info[cnt + 0:rear]
        pool.apply_async(func, (batch,))
        cnt += 20
    pool.close()
    pool.join()


if __name__ == "__main__":
    # paras
    base_dir = r'./'
    excel_demo = r'List.xlsx'
    sheet_name = 'webpage'
    # load Excel
    print('load demo Excel...')
    full_path = os.path.join(base_dir, excel_demo)
    info = load_Excel(full_path, sheet_name)

    # add Excels
    print('add Excels...')
    add_Excel(info)

    # run parser...
    print('run parser...')
    multi_processor_run(batch_processor, info)
    # finished!
    print('Program finished!')
