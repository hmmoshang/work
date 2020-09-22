import xlrd
import  pandas  as pd
from pandas import DataFrame

# 读取excel文件

filename = r'中国人寿再_20202Q_20202Q_应税.xlsx' # 文件名
data = xlrd.open_workbook(filename)
sheet = 2 # sheet 页数
col = [0,5,6,7,8,10] # 需要处理的列序列
result=[] # 总的结果
for i in col:
    col_result = []  # 每一列总结果
    if i == 0:
        for j in range(sheet):
            table = data.sheets()[j]
            col_value = table.col_values(i)
            col_presult = []  # 每一列单页结果
            for k in range(0, len(col_value), 3):
                col_presult.append(col_value[k])
            col_result += col_presult
            if j == (sheet - 1):
                result.append(col_result)
    else:
        for j in range(sheet):
            table = data.sheets()[j]
            col_value = table.col_values(i)
            col_presult = [] # 每一列单页结果
            for k in range(2,len(col_value),3):
                col_presult.append(col_value[k])
            col_result += col_presult
            if j == (sheet - 1):
                print(sum(col_result)) # 验证结果
                result.append(col_result)

# 将数字0转出为空
for i in range(1,len(result)):
    for j in range(0,len(result[i])):
        if result[i][j] == 0.0:
            result[i][j] = ''


print(result)
# 生成新的excel文件

contract = result[0] # 英文合同名称
# 账单期
period = '2020Q2'
col_period = []

col_isContract = [] # 合同/临分
# 币种
money = 'RMB'
col_money = []

premiums = result[1] # 分保费
added_tax = result[2] # 增值税
commission = result[3] # 手续费
changed_premiums = [] # 变更分保费
changed_commission = [] # 变更分保手续费
net_income = [] # 纯益手续费
refund = [] # 退保金
echanged_premiums = [] #预缴保费当期变动
amortise = result[4] # 摊赔
maturity_payment = [] #满期金
balance = result[5] # 余额
check = []
note = [] #备注

# 合同简称映射

contract_mapping = {}
mapping_table = xlrd.open_workbook( r'合同名称映射.xlsx').sheets()[0]
row_num = mapping_table.nrows
for i in range(1,row_num):
    contract_mapping[mapping_table.cell_value(i,0)] = mapping_table.cell_value(i,1)

# 去掉合同\n字段
str = '合同\n'
check_value = '-'
for i in range(0,len(contract)):
    if 'Fac' in contract[i]:
        col_isContract.append('F')
    else:
        col_isContract.append('T')
    if str in contract[i]:
        contact_name = contract[i].replace(str,'')
        contract.pop(i)
        contract.insert(i,contact_name)
print(contract) # 验证合同

# 构建数据表内容
for i in contract:
    note.append(contract_mapping[i])
    col_period.append(period)
    col_money.append(money)
    check.append(check_value)
    maturity_payment.append('')
    changed_premiums.append('')
    changed_commission.append('')
    net_income.append('')
    refund.append('')
    echanged_premiums.append('')


dic = {'合同简称': note,
        '账单期': col_period,
        '合同/临分': col_isContract,
        '币种': col_money,
        '分保费': premiums,
        '增值税': added_tax,
        '手续费': commission,
        '变更分保费': changed_premiums,
        '变更分保手续费': changed_commission,
        '纯益手续费': net_income,
        '退保金': refund,
        '预缴保费当期变动': echanged_premiums,
        '摊赔':  amortise,
        '满期金': maturity_payment,
        '余额': balance,
        'check': check,
        '备注': contract,
       }
df = pd.DataFrame(dic)
df.to_excel('应税.xlsx', index=False)


