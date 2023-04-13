from enum import Enum
from typing import Tuple
from typing import Dict
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import os
import csv
import time
import smtplib as smtp
import copy

class Message:
  """
  SMTPサーバ情報(ServerInfo)とメール情報(Email)を定義するクラス

  Attributes
  ----------
  server_info : ServerInfo
    SMTPサーバ情報を定義するオブジェクト
  email_info : EmailInfo
    メール情報を定義するオブジェクト
  """

  class ServerInfo:
    """
    SMTPサーバ情報を定義するクラス
    
    Attributes
    ----------
    address : str
      SMTPサーバのアドレス\n
      初期値は空文字
    port : int
      SMTPサーバとの通信に用いるポート番号\n
      初期値は-1
    account_address : str
      SMTPサーバとの認証で用いるメールアドレス\n
      初期値は空文字
    account_password : str
      SMTPサーバとの認証で用いるパスワード\n
      初期値は空文字
    """
    address : str = ''
    port : int = -1
    account_address : str = ''
    account_password : str = ''

  class EmailInfo:
    """
    メール情報を定義するクラス

    Attributes
    ----------
    from_address : str
      送信元メールアドレスの文字列\n
      初期値は空文字
    to_addresses : Dict[int, list[str]]
      送信順に紐づく送信先メールアドレスの辞書\n
      キーは自然数で表される送信順の整数、値は送信先メールアドレスの文字列リスト
    company_names : Dict[int, str]
      送信順に紐づく送信先会社名の辞書\n
      キーは自然数で表される送信順の整数、値は送信先会社名の文字列
    customer_names : Dict[int, str]
      送信順に紐づく送信先担当者名の辞書\n
      キーは自然数で表される送信順の整数、値は送信先担当者名の文字列
    subject : str
      メール件名の文字列
    attachment_files : list[MIMEApplication]
      添付ファイルのMIMEメッセージオブジェクトリスト
    message_content : str
      メール本文(改行含む)の文字列
    """
    from_address : str = ''
    to_addresses : Dict[int, list[str]] = {}
    company_names: Dict[int, str] = {}
    customer_names : Dict[int, str] = {}
    subject : str = ''
    attachment_files : list[MIMEApplication] = []
    message_content : str = ''
  
  server_info : ServerInfo = None
  email_info : EmailInfo = None

  def __init__(self):
    self.server_info = Message.ServerInfo()
    self.email_info = Message.EmailInfo()

class FileType(Enum):
  """
  ファイル選択ダイアログで選択するファイルの種類を定義するクラス\n
  タプルには以下の情報が含まれる\n
    1. ファイル選択ダイアログのタイトル\n
    2. 選択可能な拡張子の説明文\n
    3. 選択可能な拡張子\n
    4. 拡張子名

  Attributes
  ----------
  CSV : Tuple[str, str, str, str]
    CSVファイル
  TXT : Tuple[str, str, str, str]
    TXTファイル
  """
  CSV = ('CSVファイルを選択', 'csv files', '*.csv', 'CSV')
  TXT = ('テキストファイルを選択', 'text files', '*.txt', 'TXT')

def get_file_path(filetype: FileType) -> str:
  """
  ファイル選択ダイアログで指定したCSV, TXTファイルのデータを取得する

  Parameters
  ----------
  filetype : FileType
    クラスFileTypeで定義されたファイルの種類
  
  Returns
  -------
  file_path : str
    指定したファイルの絶対パス
  
  Raises
  ------
  ファイルが選択されなかった場合 または 指定された拡張子でなかった場合 はRuntimeError
  """
  # ファイル選択ダイアログの表示, 選択されたファイルパスの取得
  select_file_path = filedialog.askopenfilename(
    title = filetype.value[0],
    filetypes = [(filetype.value[1], filetype.value[2])], 
    initialdir = os.path.expanduser('~/Desktop')
    )

  # 指定された拡張子のファイルを選択していればそのファイルパスを代入
  if (filetype == FileType.CSV and select_file_path.endswith('.csv')) \
    or (filetype == FileType.TXT and select_file_path.endswith('.txt')):
    return select_file_path
  else:
    raise RuntimeError(filetype[3] + 'ファイルが選択されませんでした。')

def strip_list(_list: list[str]) -> list[str]:
  """
  listの各要素の前後の空白を削除したlistを返却する
  
  Parameters
  ----------
  list : list[str]
    各要素の前後に空白が含まれうる文字列リスト
  
  Returns
  -------
  各要素の前後の空白が除去された文字列リスト
  """
  return list(map(lambda e: e.strip(), _list))

def read_csv_file(message: Message, csv_file_path: str):
  """
  CSVデータをMessageオブジェクトに追加する\n
  CSVファイルは以下の内容で記述する\n
    server_address, port, email_address, password\n
    from_email_address\n
    company_name, customer_name, to_email_addresses...(複数ある場合は半角カンマ区切り)\n
    company_name, customer_name, to_email_addresses...(同上)\n
    company_name, customer_name, to_email_addresses...(同上)\n

  Parameters
  ----------
  message : Message
    SMTPサーバ情報(ServerInfo), メール情報(EmailInfo)を格納するMessageオブジェクト
  csv_file_path : str
    読込対象となるCSVファイルの絶対パスの文字列
  
  Raises
  ------
  CSVファイルが存在しない場合 または CSVデータに不備がある場合 はRuntimeError
  """
  if os.path.exists(csv_file_path):
    with open(csv_file_path) as csv_file:
      reader = csv.reader(csv_file)
      server_info_list = strip_list(next(reader))

      # 1行目 (SMTPサーバ情報)
      message.server_info.address = server_info_list[0]
      message.server_info.port = int(server_info_list[1])
      message.server_info.account_address = server_info_list[2]
      message.server_info.account_password = server_info_list[3]

      # 2行目 (送信元アドレス)
      message.email_info.from_address = ", ".join(strip_list(next(reader)))

      # 3行目以降 (会社名, 顧客名, 送信先アドレス...)
      # 会社名・顧客名はそれぞれ1社/1人しか指定できない
      to_count: int = 0
      for row in reader:
        to_count += 1
        # 空白の除去
        receiver_info_list = strip_list(row)

        if len(receiver_info_list) >= 3:
          # 1つ目の要素(送信先会社名)のポップ
          message.email_info.company_names[to_count] = receiver_info_list.pop(0)
          # 2つ目の要素(送信先顧客名)のポップ
          message.email_info.customer_names[to_count] = receiver_info_list.pop(0)
          # 残りの要素(送信先メールアドレス)の取得
          for to_email_address in receiver_info_list:
            if '@' in to_email_address:
              message.email_info.to_addresses.setdefault(to_count, []).append(to_email_address)
            else:
              raise RuntimeError((to_count + 2) + '行目のメールアドレス' + os.linesep \
                + to_email_address + os.linesep \
                + 'が不正です。')
        else:
          raise RuntimeError((to_count + 2) + '行目のCSVデータが不正です。')
  else:
    raise RuntimeError('CSVファイルが見つかりませんでした。\n' \
       + csv_file_path)

def read_text_file(email_info: Message.EmailInfo, txt_file_path: str):
  """
  TXTデータをEmailInfoオブジェクトに追加する\n
  TXTファイルは以下の内容で記述する\n
    1行目 subject\n
    2行目 attachment_file_paths(複数ある場合はカンマ区切り, 絶対パスで記載)\n
    3行目 message_contents

  Parameters
  ----------
  email_info : Message.EmailInfo
    メール情報を格納するMessage.EmailInfoオブジェクト
  txt_file_path : str
    TXTファイルの絶対パス
  
  Raises
  -------
  TXTファイルが存在しない場合 または 添付ファイルが存在しない場合 はRuntimeError
  """
  if os.path.exists(txt_file_path):
    with open(txt_file_path) as text_file:
      # 件名
      email_info.subject = text_file.readline().strip()
      
      # 添付ファイル
      attachment_file_paths: list[str] = strip_list(text_file.readline().split(','))
      for attachment_file_path in attachment_file_paths:
        # ファイルが存在する場合のみ添付する
        if os.path.exists(attachment_file_path):
          with open(attachment_file_path, 'rb') as attachment_file:
            attachment = MIMEApplication(attachment_file.read())
            attachment.add_header('Content-Disposition', 'attachment', \
              filename=os.path.basename(attachment_file_path))
            email_info.attachment_files.append(attachment)
        else:
          raise RuntimeError('以下の添付ファイルが見つかりません。' + os.linesep \
            + attachment_file_path + os.linesep \
            + 'ファイルパスを確認してください。')
        
      # メッセージ内容
      email_info.message_content = ''.join(text_file.readlines())
  else:
    raise RuntimeError('TXTファイルが見つかりませんでした。\n' \
       + txt_file_path)

def create_message(csv_file_path: str, txt_file_path: str) -> Message:
  """
  CSV, TXTデータをもとにMessageオブジェクトを生成する

  Parameters
  ----------
  csv_file_path : str
    SMTPサーバ情報、送信先情報を記載したCSVファイルパスの文字列
  txt_file_path : str
    件名、添付ファイルパス、メッセージ本文を記載したTXTファイルパスの文字列

  Returns
  -------
  CSV, TXTデータが格納されたMessageオブジェクト

  Raises
  ------
  CSV・TXTデータの読込時に発生したエラーはRuntimeError、\n
  その他の例外については例外がrethrowされる
  """
  try:
    message: Message = Message()

    # CSVデータの読込
    read_csv_file(message, csv_file_path)

    # TXTデータの読込
    read_text_file(message.email_info, txt_file_path)
    
    return message
  except RuntimeError as e:
    raise
  except:
    raise

def replace_placeholder(message_content: str, replace_dict: Dict[str, str]) -> str:
  """
  会社名と担当者名のプレースホルダを置換する

  Parameters
  ----------
  message_content : str
    置換前のメッセージ本文の文字列
  replace_dict : Dict[str, str]
    キーが置換対象の文字列、値が置換後の文字列となる辞書

  Returns
  -------
  会社名と担当者名のプレースホルダが置換されたメッセージ本文
  """
  replaced_message: str = message_content
  for key in replace_dict.keys():
    replaced_message = replaced_message.replace(key, replace_dict[key])
  
  return replaced_message

def get_error_message(exception) -> str:
  """
  例外に応じたエラーメッセージを取得する\n
  このメソッドは、例外発生時にしか呼び出してはならない

  Parameters
  ----------
  exception : Any
    処理中に発生した例外

  Returns
  -------
  例外のエラーメッセージの文字列\n
  """
  if issubclass(exception, smtp.SMTPServerDisconnected):
    return 'サーバとの接続が切断されました。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPHeloError):
    return 'サーバが HELO コマンドに応答しませんでした。サーバアドレス または ポート番号 を確認してください。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPResponseException):
    return 'SMTPエラーが発生しました。SMTPエラーコード: ' + exception.smtp_code + ', SMTPエラーメッセージ: ' + exception.smtp_error
  elif issubclass(exception, smtp.SMTPSenderRefused):
    return '送信元メールアドレスを確認してください。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPRecipientsRefused):
    return '送信先メールアドレスを確認してください。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPDataError):
    return 'メッセージ内容に問題があります。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPConnectError):
    return 'サーバへの接続時にエラーが発生しました。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPHeloError):
    return 'サーバが HELO コマンドに応答しませんでした。サーバアドレス または ポート番号 を確認してください。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPNotSupportedError):
    return 'サーバが AUTH コマンドに対応していません(SMTP認証に非対応)。エラーメッセージ: ' + exception
  elif issubclass(exception, smtp.SMTPAuthenticationError):
    return 'ユーザ名(メールアドレス) または パスワードが違います。エラーメッセージ: ' + exception
  else:
    return '何らかのエラーが発生しました。エラーメッセージを確認してください。エラーメッセージ: ' + exception

def send_message_with_sleep(sec: int, message: Message):
  """
  sec[秒]実行待機後にメールを逐次送信する

  Parameters
  ----------
  sec : int
    送信待機を行う時間[秒]
  message : Message
    SMTPサーバ情報、メール情報を保持するMessageオブジェクト
  
  Raises
  -------
  メールの作成または送信処理中に発生した例外をrethrowする
  """
  try:
    # ベースとなるMIMEオブジェクトの生成
    base_mime = MIMEMultipart()
    # 件名
    base_mime['Subject'] = message.email_info.subject
    # 送信元アドレス
    base_mime['From'] = message.email_info.from_address
    # 添付ファイル
    for attachment_file in message.email_info.attachment_files:
      base_mime.attach(attachment_file)
    
    # SMTPサーバのオブジェクト生成
    server = smtp.SMTP(message.server_info.address, message.server_info.port, timeout=10)

    if server.has_extn('STARTTLS'):
      # SSL/TLS通信の開始
      server.starttls()
    
    # メールサーバへのログイン
    server.login(message.server_info.account_address, message.server_info.account_password)

    # 個別設定
    replace_dict: dict = {}
    for key in message.email_info.company_names.keys():
      # 生成したベースMIMEオブジェクトをディープコピー
      mime = copy.deepcopy(base_mime)
      # 送信先アドレス
      mime['To'] = ', '.join(message.email_info.to_addresses[key])
      # 会社名
      replace_dict['{COMPANY_NAME}'] = message.email_info.company_names[key]
      # 顧客名
      replace_dict['{CUSTOMER_NAME}'] = message.email_info.customer_names[key]
      # メッセージ内容
      mime.attach(MIMEText(replace_placeholder(message.email_info.message_content, replace_dict)))

      # sec[秒]の送信待機(初回は送信待機しない)
      if not key == 1:
        time.sleep(sec)

      # メールの送信
      server.send_message(mime)
    
    # サーバからの切断
    server.quit()
  except:
    raise

def main():
  """
  メール送信のメイン処理
  """
  result_title: str = ''
  result_message: str = ''

  try:
    # ファイル選択ダイアログを開く
    message: Message = create_message(get_file_path(FileType.CSV), get_file_path(FileType.TXT))

    # メールの配信
    send_message_with_sleep(5, message)
  # 送信後のダイアログ生成
  except Exception as e:
    result_title = 'エラー'
    result_message = e
    messagebox.showerror(result_title, result_message)
  else:
    result_title = '送信成功'
    result_message = str(len(message.email_info.company_names.keys())) + '件すべてのメールの配信に成功しました。'
    messagebox.showinfo(result_title, result_message)