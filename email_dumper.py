from imapclient import IMAPClient
from imaplib import IMAP4
from socket import gaierror
from tqdm import tqdm
from time import time
from hashlib import md5
import os
import sys
from contextlib import contextmanager
from email.message import  Message
import smtplib
import poplib
from abc import ABC, abstractmethod
from datetime import datetime
import click


class EmailDumperInterface(ABC):
    
    @abstractmethod
    def dump_all(self,  *args, **kwargs):
        pass
    
    def print_auth_error(self):
        print('Invalid username or password. Auth is failed')
        
    def print_server_connection_error(self, server: str):
        print(f'Connection to server "{server}" is failed')

class IMAPDumper(EmailDumperInterface):

    def __init__(self, server_name: str, username: str, password: str, ssl: bool = True) -> IMAPClient:
        try:
            server: IMAPClient = IMAPClient(server_name, ssl=ssl)
            server.login(username, password)
        except IMAP4.error:
            self.print_auth_error()
            sys.exit()
        except gaierror:
            self.print_server_connection_error(server_name)
            sys.exit()
        self.client = server

    @contextmanager
    def use_folder(self, folder_name: str):
        try:
            self.client.select_folder(folder_name)
            yield
        except:
            print(f'Cannot user folder "{folder_name}"')
        finally:
            self.client.unselect_folder()

    def dump_folder(self, folder_name: str, target_folder: str) -> None:
        with self.use_folder(folder_name):
            print(f'Dump folder "{folder_name}"')
            message_ids = self.client.search()
            messages_flags = self.client.get_flags(message_ids)
            if not len(message_ids):
                return
            for message_id, flags in tqdm(messages_flags.items()):
                message = self.client.fetch([message_id], 'RFC822').get(message_id)
                self.client.set_flags([message_id], ' '.join([i.decode() for i in flags]) if len(flags) else '')
                os.makedirs(os.path.join(target_folder, folder_name), exist_ok=True)
                emial_path = os.path.join(target_folder, folder_name, md5(str(time()).encode()).hexdigest() + '.eml')
                message = message.get(b'RFC822', message)
                if not isinstance(message, bytes):
                    message = str(message).encode()
                with open(emial_path, 'wb') as f:
                    f.write(message)
        
    def dump_all(self, target_folder: str) -> None:
        for modes, delimenter, folder_name in self.client.list_folders():
            self.dump_folder(folder_name, target_folder)
        
class POP3Dumper(EmailDumperInterface):
    
    def __init__(self, server_name: str, username: str, password: str, ssl: bool = True):
        try:
            server = poplib.POP3_SSL(server_name) if ssl else poplib.POP3(server_name)
            server.user(username)
            server.pass_(password)
        except gaierror:
            self.print_server_connection_error(server_name)
            sys.exit()
        except poplib.error_proto:
            self.print_auth_error()
            sys.exit()
        self.client = server
        
    def dump_all(self, target_folder: str):
        os.makedirs(target_folder, exist_ok=True)
        file_path_pattern = os.path.join(target_folder, '%s')
        for message_id in range(len(self.client.list()[1])):
            with open(file_path_pattern % md5(str(time()).encode()).hexdigest() + '.eml', 'wb') as f:
                for line in self.client.retr(message_id + 1)[1]:
                    f.write(line + b'\n')

def send_test_message(server_name: str, username: str, password: str, to_address: str, message_text: str, ssl: bool = True):
    message = Message()
    message["From"] = username
    message["To"] = to_address
    message["Subject"] = 'Test message'
    message['Date'] = str(datetime.now())
    message.set_payload(message_text)
    server: smtplib.SMTP = smtplib.SMTP(server_name) if not ssl else smtplib.SMTP_SSL(server_name)
    server.login(username, password)
    server.auth_plain()
    server.sendmail(username, to_address, str(message))
    server.quit()

@click.group(chain=True)
@click.option('-h', '--server', type=str, required=True)
@click.option('-u', '--username', type=str, required=True)
@click.option('-P', '--password', type=str, required=True)
@click.option('-s/-ns', '--ssl/--no-ssl', default=True, required=True, show_default=True)
# @click.option('-p', '--port', type=int, )
@click.pass_context
def cli(ctx, server: str, username: str, password: str, ssl: bool):
    ctx.ensure_object(dict)
    ctx.obj['params'] = {'server_name': server, 'username': username, 'password': password, 'ssl': ssl}

@cli.command()
@click.option('-o', '--output_folder', type=str, required=True, default='./imap', show_default=True)
@click.pass_context
def imap_dump(ctx, output_folder: str):
    params = ctx.obj['params']
    dumper = IMAPDumper(**params)
    dumper.dump_all(output_folder)
    print(f'Mailbox "{params.get("username")}" dumped via imap')

@cli.command()
@click.option('-o', '--output_folder', type=str, required=True, default='./pop3', show_default=True)
@click.pass_context
def pop3_dump(ctx, output_folder: str):
    params = ctx.obj['params']
    dumper = POP3Dumper(**params)
    dumper.dump_all(output_folder)
    print(f'Mailbox "{params.get("username")}" dumped via pop3')

@cli.command()
@click.option('-t', '--to_address', type=str, required=True)
@click.option('-m', '--message_text', type=str, default='There is test message', show_default=True)
@click.pass_context
def send_message(ctx, to_address: str, message_text: str):
    params = ctx.obj['params']
    send_test_message(server_name=params.get('server_name'), username=params.get('username'),
                      password=params.get('password'), ssl=params.get('ssl'), to_address=to_address,
                      message_text=message_text)
    print(f'Test message from "{params.get("username")}" to "{to_address}" sent')

if __name__ == '__main__':
    cli(obj={})
