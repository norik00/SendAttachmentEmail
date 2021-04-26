import argparse
import glob
import os
import smtplib
import sys
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from os.path import basename

import yaml

class CustomParser(argparse.ArgumentParser):
    def error(self, message):
        print(message)
        sys.exit()


class HelpTextOperate():

    def __init__(self):
        try:
            with open('config.yml', 'r', encoding='utf-8') as yml:
                config = yaml.load(
                    stream=yml, Loader=yaml.SafeLoader
                )
                self.conf = config
        except Exception as e:
            print(e)


    def get_forward_list(self):
        """
        store name list

        Returns:
            stores list: yaml
        """

        branches = [branch for branch in self.conf['branches'].keys()]
        stores = self.conf['stores']

        forward_list = {}

        for branch in branches:
            for store in stores:
                if branch == stores[store]['branch']:
                    if branch in forward_list:
                        forward_list[branch].append(store)
                    else:
                        forward_list[branch] = [store]
                else:
                    continue

        return yaml.dump(forward_list, allow_unicode=True)


class SendEmailOperate():

    def __init__(self, args):
        """
        Load config.yml file.
        Get argeuments
        """

        try:
            with open('config.yml', 'r', encoding='utf-8') as yml:
                config = yaml.load(
                    stream=yml, Loader=yaml.SafeLoader
                )
        except Exception as e:
            print(e)
            sys.exit()
        else:
            self.conf = config
            self.server = smtplib.SMTP(self.conf['server'], self.conf['port'])
            self.source = self.conf['source']
            self.cc = self.conf['cc']
            self.body = self.conf['body']
            
            # SS名あるか確認
            if args.fwdest not in self.conf['stores']:
                print(f"There is no store name in config.yaml: '{args.fwdest}' help command: '-h'\n")
                sys.exit()

            # 引数取得
            if args.division == 'car':
                self.division_ja = '車両'
            elif args.division == 'profit':
                self.division_ja = '利益'

            self.name = args.name
            self.fw_dest = args.fwdest
            self.month = args.month


    def get_to(self):
        """
        Get mail address of TO

        Return: string comma
        """
        
        # 支店名取得
        branch = self.conf['stores'][self.fw_dest]['branch']
        
        if branch in self.conf['branches']:
            # 支店[branch]のアドレス[mail]取得
            mail = self.conf['branches'][branch]['mail']

            # 店舗[sotore]のアドレス[mail]取得
            mail.extend(self.conf['stores'][self.fw_dest]['mail'])

            return ','.join(mail)

        mail = self.conf['stores'][self.fw_dest]['mail']

        return ','.join(mail)

    
    def get_cc(self):
        """
        Get mail address of CC

        Return: 
            string camma
        """

        # 支店名取得
        branch = self.conf['stores'][self.fw_dest]['branch']
        mail = self.cc
        
        if branch not in self.conf['others']:
            return ','.join(mail)

        if branch in self.conf['others']:
            mail.extend(self.conf['others'][branch]['mail'])

            return ','.join(mail)

    
    def get_to_name(self):
        """
        Get name of TO

        Parameter:
            to_forward: string
        
        Return: 
            list <br> separate
        """

        # 支店名取得
        branch = self.conf['stores'][self.fw_dest]['branch']

        if branch in self.conf['branches']:

            if self.conf['branches'][branch]['order'] == 'last':
                addressee = self.conf['stores'][self.fw_dest]['addressee']
                addressee.extend(self.conf['branches'][branch]['addressee'])            
            else:
                addressee = self.conf['branches'][branch]['addressee']
                addressee.extend(self.conf['stores'][self.fw_dest]['addressee'])
            
            return '<br>'.join(addressee)

        addressee = self.conf['stores'][self.fw_dest]['addressee']

        return '<br>'.join(addressee)
    

    def get_message(self):
        """
        Get message
            
        Return:
            MIMEMultipart
        """

        addressee = self.get_to_name()
        message = self.body.format(
            addressee=addressee, division_ja=self.division_ja
        )

        msg = MIMEMultipart()
        msg['Date'] = formatdate()

        # Message header.
        # For auto filtering of email client.
        msg['Header'] = Header('TRANSFER')

        msg["Subject"] = Header(f'【{self.month}月】{self.division_ja}転送金額（{self.fw_dest}　{self.name}）')
        msg["To"] = self.get_to()
        msg["CC"] = self.get_cc()
        msg["From"] = self.source

        # Save message in email client too.
        msg["Bcc"] = self.source

        msg.attach(MIMEText(message, 'html'))
        
        return msg

    
    def attach_file(self, msg):
        """
        Attach file
        添付ファイル名形式： 【m月】{division_ja}転送金額（{fw_dest} {name}）

        Parameter:
            msg: MIMEMultipart
            
        Return:
            tupple:
                boolean:
                msg: MIMEMultipart or ''
        """

        file_name = f'【{self.month}月】{self.division_ja}転送金額（{self.fw_dest}　{self.name}）.xlsx'

        # ディレクトリ内の全excel(xslx）を取得
        files = glob.glob(self.conf['directory'] + '/*.xlsx')

        for file_path in files:
            if basename(file_path) == file_name:
                with open(file_path, "rb") as f:
                    part = MIMEApplication(
                        f.read(),
                        Name=basename(file_path)
                    )
        
                    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file_path)
                    msg.attach(part)

                    return True, msg
            
        return False, ''


    def send_mail(self, msg):
        """
        Send email

        Parameters:
            msg: MIMEMultipart
        
        Return:
            tupple:
                boolean:
                string: '' or error
        """
        
        try:
            self.server.send_message(msg)
            self.server.quit()

            print(f"Completed. 転送月: {self.month}, 転送先: {self.fw_dest}, 車名: {self.name}")

            return True, ''
        except Exception as e:
            print('send email is False.')

            return False,e


class CheckMessageOperate():
    """
    Before send email.

    Verification
    """

    def get_subject(self, msg):
        charset = msg.get_content_charset()
        payload = msg['Subject']
        try:
            if payload:
                if charset:
                    return payload.decode(charset)
                else:
                    return payload.decode()
            else:
                return ""
        except Exception:
            return payload


    def get_content(self, msg):
        for part in msg.walk():
            payload = part.get_payload(decode=True)

            if payload is None:
                continue
            
            charset = part.get_content_charset()
            
            if charset:
                return payload.decode(charset)
            else:
                return payload.decode()

        return ''


# if __name__ == '__main__':
def main():
    help_operate = HelpTextOperate()

    # argparseの設定
    parser = CustomParser(
        description='Auto sending email',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
attachment file:
name format    【[month]】[car or profit]（[forwarding destination]　[vehicle name]）

forward list:  
{help_operate.get_forward_list()}
"""
    )

    ## 必須の引数
    parser.add_argument(
        '-d','--division',
        help='forwarding division',
        default='car',
        choices=['car', 'profit'],
        required=True
    )
    parser.add_argument(
        '-n','--name',
        help='vehicle name',
        required=True
    )
    parser.add_argument(
        '-fd', '--fwdest',
        help='forwarding destination',
        required=True
    )
    parser.add_argument(
        '-m', '--month',
        help='month',
        required=True
    )

    # 引数取得
    args = parser.parse_args()
    
    operate = SendEmailOperate(args)

    message = operate.get_message()

    # メッセージ内容確認
    msg_operate = CheckMessageOperate()
    subject = msg_operate.get_subject(message)
    body = msg_operate.get_content(message).replace('<br>', '\r\n')
    
    print(f"\r\n")
    print("===========================================================================")
    
    print(f"TO: {message['To']}")
    print(f"CC: {message['CC']}")
    print(f'Subject: {subject}')
    print(f'Body:\r\n{body}')

    print("===========================================================================")
    print(f"\r\n")

    is_valid = input('Do you wish submit the mail? Y/n >> ')

    if is_valid.lower() == 'y':
        is_file, mail_message = operate.attach_file(message)
        
        if is_file:
            is_send, error = operate.send_mail(mail_message)
            if is_send:
                print('Completed')
            else:
                print(error)
        else:
            print('There is no attachment. Check in directory.')
    else:
        print('Exit')

