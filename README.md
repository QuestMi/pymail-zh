# Pymail-zh

![Contributions welcome](https://img.shields.io/badge/contributions-welcome-orange.svg)

`Pymail-zh` 基于 [IMAPClient](https://github.com/mjs/imapclient) 封装了一些中文邮件的处理，以及一些 helper function。使用者可以轻松的拿到邮件的结构化数据。
同时，基于不同邮箱 imap server 的实现，封装了一些特殊的方法，例如移动邮件。

Mail 对象含以下属性：

```python
uid  # 邮件 uid
eml  # 邮件原始数据
sender  # 邮件发送人
subject  # 邮件主题
body  # 邮件主体
mail_date  # 邮件日期
attachments  # 邮件附件
```

## 安装

```bash
pip install pymail-zh
```

## 使用

获取邮件，移动邮件到指定目录（目录不存在自动创建），保存邮件到本地 eml。

```python
import pathlib

from pymail_zh.mail_client import MailQQ, MailQiYeQQ, MailQiYeAliyun, MailAliyun


def filter_mail_subject_move_and_save_eml(mail_client, mail):
    """
    移动邮件，并保存到本地。
    :param mail_client: MailQiYeQQ 等的实例
    :param mail: Mail 对象
    :return:
    """
    if '测试' in mail.subject:
        mail_client.move_mail(mail.uid, 'Parent_1/Parent_2')
    with open(pathlib.Path(__file__).parent / f"{mail.subject}.eml", 'wb') as f:
        f.write(mail.eml)


with MailQiYeQQ(user_name='YOUR_EMAIL', password='MAIL_PASSWORD') as qq_mail_client:
    # UID 不是每次都一样，所以拿到邮件的 UID 时，需要交给 callback 处理。
    # 单次收取邮件数目不超过30个，有的邮箱拿多了会 ban。
    qq_mail_client.handle_mails(filter_mail_subject_move_and_save_eml, mails_count=30)
```

读取 eml 文件，获取邮件结构化信息

```python
from pymail_zh.mail_utils import eml_to_mail_info

with open('转发：测试邮件1.eml', 'rb') as mail_file:
    eml_to_mail_info(mail_file.read())
```