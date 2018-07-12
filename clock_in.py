from openpyxl import load_workbook
import csv


def get_map(path):
    wechat2name = {}
    qq2name = {}
    workbook = load_workbook(path)
    sheets = workbook.sheetnames
    booksheet = workbook[sheets[0]]
    rows = booksheet.rows
    for row in rows:
        line = [col.value for col in row]
        if line[1] and line[1].startswith("o86"):
            wechat2name[line[1]] = line[2]
        if line[0] and isinstance(line[0],int):
            qq2name[line[0]] = line[2]
    return wechat2name, qq2name


def check_talbe(path,wechat2name,qq2name):
    names = {'A':set(),'B':set(),'C':set()}
    with open(path,'r') as fin, open("clock_in.log","w") as log:
        rows = csv.reader(fin, delimiter=',')
        next(rows)
        for line in rows:
            print(line)
            if len(line[7]) > 1:
                if len(line[7]) < 4:
                    names[line[6][0]].add(line[7])
                else:
                    print("name:", line[7], file=log)
            if line[4].startswith("o86"):
                if line[4] in wechat2name:
                    name = wechat2name[line[4]]
                    names[line[6][0]].add(name)
                else:
                    print("wechat:", line[4], file=log)
            if len(line[3]) > 1:
                qq = int(line[3])
                if qq in qq2name:
                    name = qq2name[qq]
                    names[line[6][0]].add(name)
                else:
                    print("qq:", qq, file=log)
    return names


def clock_in(path, names):
    workbook = load_workbook(path)
    sheets = workbook.sheetnames
    booksheet = workbook[sheets[1]]
    col = next(booksheet.columns)
    name_list = [cell.value for cell in col]
    name = set(name_list[1:])
    for col in booksheet.columns:
        pass
    for i, each in enumerate(col):
        if each.value == "已回家":
            names['C'].add(name_list[i])
    print("不顺利",names['B'])
    print("已回家",names['C'])
    print("未签到",name.difference(names['A'].union(names['B']).union(names['C'])))


if __name__ == "__main__":
    w, q = get_map('数据表.xlsx')
    n = check_talbe('2256106_seg_1.csv', w, q)
    clock_in('2018暑期住宿.xlsx', n)
