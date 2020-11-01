# python操作mysql数据库
import pymysql
import datetime
from openpyxl import load_workbook


class DutyDB:

    def open(self):
        try:
            self.con = pymysql.connect(
                host='127.0.0.1',
                port=3306,
                user='root',
                password='root',
                database='jDB',
                charset='utf8')
            print('Connection Database Successful！')
            self.cursor = self.con.cursor(pymysql.cursors.DictCursor)
        except Exception as con_err:
            print(con_err)

        try:
            sql = 'create table duty (' \
                  'date DATE primary key,' \
                  'ck FLOAT, ' \
                  'hl FLOAT, ' \
                  'idc FLOAT, ' \
                  'user_BW FLOAT, ' \
                  'user_ALL FLOAT)'
            self.cursor.execute(sql)
        except Exception as cre_err:
            print(cre_err)

    def save(self):
        self.con.commit()
        print('save successful and continue...')

    def exit(self):
        self.con.commit()
        self.con.close()
        print('exit and save successful')

    def exit_no_save(self):
        self.con.close()
        print('exit and not save')

    def clear(self):
        try:
            self.cursor.execute('delete from duty')
            print('delete all data successful')
        except Exception as del_err:
            print(del_err)

    def show(self):
        try:
            self.cursor.execute('select * from duty')
            print('%s\t\t\t%s\t\t%s\t\t%s\t\t%s\t\t%s' % ('日期', '出口', '互联', 'IDC', '本网用户', '总用户'))
            rows = self.cursor.fetchall()
            for row in rows:
                print('%s\t%s\t%s\t%s\t%s\t\t%s' % (row['date'], row['ck'], row['hl'], row['idc'], row['user_BW'], row['user_ALL']))
        except Exception as show_err:
            print(show_err)

    def __insert(self, date, ck, hl, idc, user_BW, user_ALL):
        try:
            sql = 'insert into duty (date, ck, hl, idc, user_BW, user_ALL) values (%s, %s, %s, %s, %s, %s)'
            self.cursor.execute(sql, (date, ck, hl, idc, user_BW, user_ALL))
            print(self.cursor.rowcount, 'row inserted')
        except Exception as insert_err:
            print(insert_err)

    def __update(self, date, ck, hl, idc, user_BW, user_ALL):
        try:
            sql = 'update duty set ck=%s, hl=%s, idc=%s, user_BW=%s, user_ALL=%s where date=%s'
            self.cursor.execute(sql, (ck, hl, idc, user_BW, user_ALL, date))
            print(self.cursor.rowcount, 'row updated')
        except Exception as up_err:
            print(up_err)

    def __delete(self, date):
        try:
            sql = 'delete from duty where date=%s'
            self.cursor.execute(sql, (date,))
            print(self.cursor.rowcount, 'row deleted')
        except Exception as del_err:
            print(del_err)

    def enterduty(self, s):
        while True:
            m = input(s)
            try:
                m = float(m)
                break
            except Exception as err:
                print(err)
        return m

    def insert(self):
        date = input('日期：').strip()
        if date != '':
            ck = self.enterduty('出口：')
            hl = self.enterduty('互联：')
            idc = self.enterduty('IDC：')
            user_BW = self.enterduty('本网用户：')
            user_ALL = self.enterduty('总用户：')
            self.__insert(date, ck, hl, idc, user_BW, user_ALL)
        else:
            print('日期不能为空')

    def update(self):
        date = input('日期：').strip()
        if date != '':
            ck = self.enterduty('出口：')
            hl = self.enterduty('互联：')
            idc = self.enterduty('IDC：')
            user_BW = self.enterduty('本网用户：')
            user_ALL = self.enterduty('总用户：')
            self.__update(date, ck, hl, idc, user_BW, user_ALL)
        else:
            print('日期不能为空')

    def delete(self):
        date = input('日期：').strip()
        if date != '':
            self.__delete(date)
        else:
            print('日期不能为空')

    def export(self):
        x = input('导出日期？（格式2020-07-16；默认前一天）：')
        if x =='':
            export_day = (datetime.datetime.now() -  datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            export_riqi = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y年\n%m月\n%d日')
        else:
            export_day = x
            export_riqi = x[:4] + '年\n' + x[5:7] + '月\n' + x[8:10] + '日'

        try:
            sql = 'select * from duty where date = %s'
            self.cursor.execute(sql, (export_day,))
            row = self.cursor.fetchone()

            excel_file = '每日监控互联网宽带业务流量流向数据分析表.xlsx'
            workbook = load_workbook(filename=excel_file)
            sheet = workbook.active

            cell_riqi = sheet['A1']
            cell_riqi.value = export_riqi
            cell_chukou = sheet['C3']
            cell_chukou.value = str(row['ck']) + ' Gbps'
            cell_hulian = sheet['C4']
            cell_hulian.value = str(row['hl']) + ' Gbps'
            cell_idc = sheet['C5']
            cell_idc.value = str(row['idc']) + ' Gbps'
            cell_benwang = sheet['C7']
            cell_benwang.value = str(row['user_BW']) + ' 万'
            cell_zongshu = sheet['C8']
            cell_zongshu.value = str(row['user_ALL']) + ' 万'

            cell_zongliuliang = sheet['C6']
            cell_zongliuliang.value = str('%.2f' % (row['ck'] + row['hl'] + row['idc'])) + ' Gbps'
            cell_bili_ck = sheet['D3']
            cell_bili_ck.value = str('%.2f' % ((row['ck']/(row['ck'] + row['hl'] + row['idc']))*100)) + '%'
            cell_bili_hl = sheet['D4']
            cell_bili_hl.value = str('%.2f' % ((row['hl']/(row['ck'] + row['hl'] + row['idc']))*100)) + '%'
            cell_bili_idc = sheet['D5']
            cell_bili_idc.value = str('%.2f' % ((row['idc'] / (row['ck'] + row['hl'] + row['idc'])) * 100)) + '%'
            cell_pengboshi = sheet['E7']
            cell_pengboshi.value = str('%.2f' % (row['user_ALL'] - row['user_BW'])) + ' 万'
            cell_hujun = sheet['C9']
            cell_hujun.value = str('%.2f' % (((row['ck'] + row['hl'] + row['idc'])*100)/row['user_BW'])) + ' Kbps'
            cell_neiwanglv = sheet['C10']
            cell_neiwanglv.value = str('%.2f' % (((row['hl'] + row['idc'])/(row['ck'] + row['hl'] + row['idc']))*100)) + '%'

            workbook.save(filename=excel_file)
            print('Excel表格导出完毕！')
        except Exception as ex_err:
            print(ex_err)

    def process(self):
        self.open()
        while True:
            s = input('>')
            if s == 'show':
                self.show()
            elif s == 'insert':
                self.insert()
            elif s == 'update':
                self.update()
            elif s == 'delete':
                self.delete()
            elif s == 'export':
                self.export()
            elif s == 'clear':
                self.clear()
            elif s == 'save':
                self.save()
            elif s == 'exit':
                self.exit()
                break
            elif s == 'exit_no_save':
                self.exit_no_save()
                break
            else:
                print('Accept show/insert/update/delete/export/clear/exit')
                print('show --- show the rows')
                print('insert --- insert a new row')
                print('update --- update a row')
                print('delete --- delete a row')
                print('export --- exprot to Duty.xlsx')
                print('clear --- clear all data')
                print('save --- save and continue...')
                print('exit --- exit and save')
                print('exit_no_save --- exit and not save')


db = DutyDB()
db.process()


