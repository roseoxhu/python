#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys
import logging

# pip install --trusted-host pypi.douban.com numpy pandas scipy matplotlib
import pandas as pd
# pandas里最常用的和Excel I/O有关的四个函数是read_csv/ read_excel/ to_csv/ to_excel
# https://pyecharts.org/#/zh-cn/
# pip install --trusted-host pypi.douban.com pyecharts

# excel文件和pandas的交互读写，主要使用到pandas中的两个函数,
# 一个是pd.ExcelFile函数,一个是DataFrame.to_excel函数
# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html

'''
微信: 商品名称,商户订单号,总金额
支付宝: 商品名称,商户订单号,订单金额（元）
销售订单: 订单号,交易类型,拆单类型,支付流水号,在线支付类型,在线支付金额
    交易类型: 交易/退款
    拆单类型: 未拆单,子单
    在线支付类型: 微信*/支付宝*
按 商户订单号=支付流水号 分组比对
'''

def round_up(value):
    '将数值放大100倍，精确的四舍五入'
    # 替换内置round函数,实现保留2位小数的精确四舍五入
    return round(value * 100) / 100.0

def parse_sales_flow_wechatpay(file_name):
    '解析销售订单微信交易流水'
    sheet_name = 'SheetJS'
    header = 0 # Row (0-indexed) to use for the column labels of the parsed DataFrame.
    col_names = 'C,D,F,O,P,Q' # 订单号,交易类型,拆单类型,支付流水号,在线支付类型,在线支付金额
    df = pd.read_excel(file_name, sheet_name, header=header, usecols=col_names, 
        converters={'支付流水号':str.strip,'在线支付金额':float})
    print(df.columns)
    #使用“与”条件进行筛选 微信销售订单交易流水
    # wechatpay_df = df[df['在线支付类型'].apply(lambda payType: str(payType).startswith('微信'))] 
    wechatpay_df = df[df['在线支付类型'].str.contains('微信') 
                    & (df['交易类型'] == '交易') & (df['拆单类型'] != '母单')]
    print(wechatpay_df)
    logging.debug('行数: %d' % len(wechatpay_df))
    logging.debug('累计总金额: %.2f' % wechatpay_df['在线支付金额'].sum())

    # 分组汇总
    df_group_sum = wechatpay_df.groupby(['支付流水号'])[['在线支付金额']].agg('sum')
    print(df_group_sum)
    logging.debug('行数: %d' % len(df_group_sum))
    logging.debug('累计总金额: %.2f' % df_group_sum['在线支付金额'].sum())
    # 添加累计汇总行
    df_group_sum.loc['累计总金额'] = df_group_sum['在线支付金额'].sum()
    return df_group_sum


def parse_finance_flow_wechatpay(file_name):
    '解析微信交易流水'
    sheet_name = '微信'
    header = 4 # Row (0-indexed) to use for the column labels of the parsed DataFrame.
    # col_names = 'Y,Z,AA' # 商户订单号=支付流水号 （老版本工作正常）
    col_names = ['商品名称','商户订单号','总金额']
    caretConv = lambda x: str(x).replace('`', '') 
    # Starting with Pandas version 2.0 all arguments of read_excel except for the first 2 arguments will be keyword-only
    # 注意：重复的字段名，Pandas 新版需要明确指定 mangle_dupe_cols=False
    # Duplicate columns will be specified as 'X', 'X.1', ...'X.N'
    df = pd.read_excel(file_name, sheet_name, header=header, usecols=col_names, 
        converters={'商户订单号': caretConv,'商品名称': caretConv})
    print(df) # 新版使用'Y,Z,AA'列名，结果重复字段：商品名称.1   商户订单号.1   总金额.1
    # type(df['商户订单号']) <class 'pandas.core.series.Series'>
    # df['商户订单号'].duplicated()
    # df['商户订单号长度'] = df['商户订单号'].apply(lambda ordernum: len(ordernum)
    # 排除微信小程序支付订单
    df = df[ df['商户订单号'].apply(lambda ordernum: len(ordernum) <= 17) ]
    print(df) # df[ df['商户订单号长度'] <= 17 ]
    logging.debug('行数: %d' % len(df))
    logging.debug('累计总金额: %s' % df['总金额'].sum())

    # 分组汇总
    df_group_sum = df.groupby(['商户订单号'])[['总金额']].agg('sum')
    print(df_group_sum)
    logging.debug('行数: %d' % len(df_group_sum))
    logging.debug('累计总金额: %.2f' % df_group_sum['总金额'].sum())

    # 数据透视表
    # df_group_sum = pd.pivot_table(df2, values=['总金额'], index=['商户订单号'], aggfunc='sum')
    # print(df_group_sum)
    # logging.debug('行数: %d' % len(df_group_sum))
    # logging.debug('累计总金额: %s' % df_group_sum['总金额'].sum())
    
    # 输出汇总
    # (file, ext) = os.path.splitext(file_name)
    # sum_file_name = "%s-汇总%s" % (file, ext)
    # logging.debug('正在写入汇总文件: %s, sheet=%s' % (sum_file_name, sheet_name))
    # df_group_sum.to_excel(sum_file_name, sheet_name)
    # logging.debug('完成写入汇总文件')
    # 添加累计汇总行
    df_group_sum.loc['累计总金额'] = df_group_sum['总金额'].sum()
    return df_group_sum


def parse_sales_flow_alipay(file_name):
    '解析销售订单支付宝交易流水'
    sheet_name = 'SheetJS'
    header = 0 # Row (0-indexed) to use for the column labels of the parsed DataFrame.
    col_names = 'C,D,F,O,P,Q' # 订单号,交易类型,拆单类型,支付流水号,在线支付类型,在线支付金额
    df = pd.read_excel(file_name, sheet_name, header=header, usecols=col_names, 
        converters={'支付流水号':str.strip,'在线支付金额':float})
    print(df.columns)
    #使用“与”条件进行筛选 支付宝销售订单交易流水
    # alipay_df = df[df['在线支付类型'].apply(lambda payType: str(payType).startswith('支付宝'))] 
    alipay_df = df[df['在线支付类型'].str.contains('支付宝') 
                    & (df['交易类型'] == '交易') & (df['拆单类型'] != '母单')]
    print(alipay_df)
    logging.debug('行数: %d' % len(alipay_df))
    logging.debug('累计总金额: %.2f' % alipay_df['在线支付金额'].sum())
    
    # 分组汇总
    df_group_sum = alipay_df.groupby(['支付流水号'])[['在线支付金额']].agg('sum')
    print(df_group_sum)
    logging.debug('行数: %d' % len(df_group_sum))
    logging.debug('累计总金额: %.2f' % df_group_sum['在线支付金额'].sum())
    # 添加累计汇总行
    df_group_sum.loc['累计总金额'] = df_group_sum['在线支付金额'].sum()
    return df_group_sum


def parse_finance_flow_alipay(file_name):
    '解析支付宝交易流水'
    sheet_name = '支付宝'
    header = 4 # Row (0-indexed) to use for the column labels of the parsed DataFrame.
    col_names = 'B,D,L' # 商户订单号=支付流水号
    # Starting with Pandas version 2.0 all arguments of read_excel except for the first 2 arguments will be keyword-only
    df = pd.read_excel(file_name, sheet_name, header=header, usecols=col_names, 
        skipfooter=1, converters={'商户订单号':str.strip})
    print(df.columns)
    print(df)
    logging.debug('行数: %d' % len(df))
    logging.debug('累计总金额: %s' % df['订单金额（元）'].sum())

    # 分组汇总
    df_group_sum = df.groupby(['商户订单号'])[['订单金额（元）']].agg('sum')
    print(df_group_sum)
    logging.debug('行数: %d' % len(df_group_sum))
    total_amount = df_group_sum['订单金额（元）'].sum()
    logging.debug('累计总金额: %.2f' % total_amount)

    # 添加累计汇总行
    df_group_sum.loc['累计总金额'] = df_group_sum['订单金额（元）'].sum()
    return df_group_sum


def write_to_excel(file_name, dataFrameMap):
    '输出汇总文件'
    writer = pd.ExcelWriter(file_name)
    for sheet_name, df in dataFrameMap.items():
        logging.debug('正在写入汇总文件: %s, sheet=%s' % (file_name, sheet_name))
        df.to_excel(excel_writer=writer, sheet_name=sheet_name)
        writer.save()
    writer.close()
    logging.debug('完成写入汇总文件')


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Usage: %s <财务流水.xlsx> <销售订单日报.xlsx>' % os.path.abspath(sys.argv[0]))
        sys.exit(0)

    logging.basicConfig(level=logging.NOTSET, 
        format='%(asctime)s - %(filename)s[line:%(lineno)d/%(thread)d] - %(levelname)s: %(message)s') # 设置日志级别
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    logging.debug("os.getcwd() = %s" % os.getcwd())

    # 支付交易流水
    # file_name = r'/path/to/财务流水.xlsx'
    file_name = sys.argv[1]
    logging.debug("xls name: %s" % file_name)
    wechatpay_df = parse_finance_flow_wechatpay(file_name)
    alipay_df = parse_finance_flow_alipay(file_name)

    # 输出汇总
    (file, ext) = os.path.splitext(file_name)
    sum_file_name = "%s-汇总%s" % (file, ext)
    write_to_excel(sum_file_name, {r'微信': wechatpay_df, r'支付宝': alipay_df})
    
    # 销售交易流水
    # file_name = r'/path/to/销售订单日报_202010.xlsx'
    file_name = sys.argv[2]
    logging.debug("xls name: %s" % file_name)
    sale_wechatpay_df = parse_sales_flow_wechatpay(file_name)
    sale_alipay_df = parse_sales_flow_alipay(file_name)
    # 输出汇总
    (file, ext) = os.path.splitext(file_name)
    sum_file_name = "%s-汇总%s" % (file, ext)
    # write_to_excel(sum_file_name, {r'微信': sale_wechatpay_df, r'支付宝': sale_alipay_df})

    # 按交易流水号连接比较差异
    join_wechatpay_df = wechatpay_df.join(sale_wechatpay_df, how="outer")
    join_wechatpay_df.fillna(value={'总金额': 0, '在线支付金额': 0}, inplace=True)
    join_wechatpay_df['差异'] = round(join_wechatpay_df['总金额'] - join_wechatpay_df['在线支付金额'], 2)
    print(join_wechatpay_df[join_wechatpay_df['差异'] != 0.0])

    join_alipay_df = alipay_df.join(sale_alipay_df, how="outer")
    join_alipay_df.fillna(value={'订单金额（元）': 0, '在线支付金额': 0}, inplace=True)
    join_alipay_df['差异'] = round(join_alipay_df['订单金额（元）'] - join_alipay_df['在线支付金额'], 2)
    print(join_alipay_df[join_alipay_df['差异'] != 0.0])
    write_to_excel(sum_file_name, {r'微信': sale_wechatpay_df, r'支付宝': sale_alipay_df, 
        r'微信-差异': join_wechatpay_df, r'支付宝-差异': join_alipay_df})

    # writer = pd.ExcelWriter(sum_file_name)
    # sheet_name='支付宝'
    # logging.debug('正在写入汇总文件: %s, sheet=%s' % (sum_file_name, sheet_name))
    # # alipay_df['商户订单号'] = alipay_df['商户订单号'].astype(str) # 转为字符串，否则按科学记数法显示
    # alipay_df.to_excel(excel_writer=writer, sheet_name='支付宝')
    # writer.save()
    # writer.close()
    # logging.debug('完成写入汇总文件')

    #xls_file = pd.ExcelFile(file_name)
    # print(xls_file.sheet_names) #显示出读入excel文件中的表名字
    #logging.debug("xls sheet: %s" % xls_file.sheet_names) # '微信', '支付宝', 'Sheet3']
