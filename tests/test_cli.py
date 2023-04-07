import unittest
from broadcastemail.cli import read_csv_file
from broadcastemail.cli import read_text_file
from private import Private
from broadcastemail.cli import Message
import pdb
from urllib.parse import unquote

class CliTestCase(unittest.TestCase):
  def test_read_csv_file(self):
    message: Message = Message()
    read_csv_file(message, Private.csv_file_path)

    self.assertEqual(message.server_info.address, Private.server_address)
    self.assertEqual(message.server_info.port, Private.server_port)
    self.assertEqual(message.server_info.account_address, Private.email_address)
    self.assertEqual(message.server_info.account_password, Private.password)
    self.assertEqual(message.email_info.company_names[1], 'あいう株式会社')
    self.assertEqual(message.email_info.customer_names[1], 'あいう　太郎')
    self.assertEqual(message.email_info.to_addresses[1], ['aiueo_test@test.test'])
    self.assertEqual(message.email_info.to_addresses[2], ['k1_test@test.test', 'k2_test@test.test', 'k3_test@test.test'])
    self.assertEqual(message.email_info.customer_names[2], 'かきく 次郎')
    self.assertEqual(message.email_info.customer_names[3], 'さしす三郎')
  
  def test_read_txt_file(self):
    message: Message = Message()
    read_text_file(message.email_info, Private.txt_file_path)

    self.assertEqual(message.email_info.subject, '件名テスト')
    self.assertEqual(unquote(message.email_info.attachment_files[0].items()[3][1].split('\'')[-1]), '添付test_1.txt')
    self.assertEqual(unquote(message.email_info.attachment_files[1].items()[3][1].split('\'')[-1]), '添付test_2.txt')
    self.assertEqual(unquote(message.email_info.attachment_files[2].items()[3][1].split('\'')[-1]), '添付test_3.txt')

    message_content: str = """これは1行目です。下に空白行が1行あります。

これは3行目です。下に空白行が2行あります。


これは6行目です。下に空白行はありません。"""
    self.assertEqual(message.email_info.message_content, message_content)