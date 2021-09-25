# ==================
# === 群发邮件脚本 ===
# ==================
# 参考：
# https://zhuanlan.zhihu.com/p/269052889
# https://zhuanlan.zhihu.com/p/109551738
# https://github.com/milesial/Pytorch-UNet

import smtplib
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
import pandas as pd
import argparse
import os
import pathlib


# ===
# === 设置发送人、发送昵称、密码、邮箱服务器
# ===
FROM_ADDR = '1097985743@qq.com'
FROM_ALIAS = '于松黎'
PASSWORD = 'abcdefghijklmnopqrst'
SERVER = smtplib.SMTP_SSL('smtp.qq.com', 465)

# ===
# === 设置邮件列表文件(要求是xlsx格式，并且表格包含一列是'to'，且包含一列是'cc')
# ===
MAIL_LIST = os.path.join('assets', 'list.xlsx')


# ======================================================================================================================


def get_args():
    parser = argparse.ArgumentParser(description='Mail sender')
    parser.add_argument('--subject', '-s', dest='mail_subject', type=str, help='Mail subject to send')
    parser.add_argument('--content', '-c', dest='txt_filename', type=str, help='Txt file for mail content')
    parser.add_argument('--attachment', '-a', dest='attachment_filename', type=str, help='Attachment')
    return parser.parse_args()


def read_list():
    df = pd.read_excel(MAIL_LIST, sheet_name='Sheet1')
    # 收件人
    df_to = df['to'].values.tolist()
    # 抄送人
    df_cc = df['cc'].values.tolist()

    # 设置接收人
    to_addrs = df_to
    if not type(df_cc[0]) == float:
        to_addrs += df_cc
    return df_to, df_cc, to_addrs


def send(mail_subject, txt_filename, attachment_filename, df_to, df_cc, to_addrs):
    # 设置邮件头
    message = MIMEMultipart()
    message['From'] = formataddr([FROM_ALIAS, FROM_ADDR])
    message['Subject'] = mail_subject
    message['To'] = ';'.join(df_to)
    if not type(df_cc[0]) == float:
        message['Cc'] = ';'.join(df_cc)

    # 读取txt内容到变量
    with open(txt_filename, 'r', encoding='utf-8') as txt:
        txt_content = txt.read()

    # 添加正文
    add_body = MIMEText(txt_content, _subtype='plain', _charset='utf-8')
    message.attach(add_body)

    # 构建邮件附件
    part_attach1 = MIMEApplication(open(attachment_filename, 'rb').read())  # 打开附件
    part_attach1.add_header('Content-Disposition', 'attachment', filename=pathlib.Path(attachment_filename).name)  # 为附件命名
    message.attach(part_attach1)  # 添加附件

    try:
        SERVER.login(FROM_ADDR, PASSWORD)
        SERVER.sendmail(FROM_ADDR, to_addrs, message.as_string())
        print('success')
        SERVER.quit()
    except smtplib.SMTPException as e:
        print('error', e)


if __name__ == '__main__':
    args = get_args()
    df_to, df_cc, to_addrs = read_list()
    send(args.mail_subject, args.txt_filename, args.attachment_filename, df_to, df_cc, to_addrs)
