from ast import arg
import socket
import json
from threading import Thread
import os

def get_file_contents(file_name):
    #print("get_file_contents:" + file_name)
    with open(file_name, 'r', encoding='utf-8') as f:
        return f.read()
    
class TcpServer:
    """
    Our actual HTTP server which will keep a long conection.
    """
    def __init__(self):
        config_info = get_file_contents(r'config/config-tcp.json')
        info_json = json.loads(config_info)
        host = info_json["ip"]
        port = info_json["port"]
    
        print(f"TCP Server started. Listening at http://{host}:{port}/")
        self.host = host
        self.port = port
        self.working_dir = os.getcwd()
        self.working_dir += '\\'
        print("workpath:" + self.working_dir)        
        
    def startSever(self):
        print("start listen")
        self.setup_socket()
        self.accept()
        self.teardown_socket()

    def setup_socket(self):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind((self.host, self.port))
        self.sock.listen(128)
        

    def teardown_socket(self):
        if self.sock is not None:
            self.sock.shutdown()
            self.sock.close()

    def accept(self):
        #print("-------accept")
        while True:
            (client, address) = self.sock.accept()
            self.client = client            
            th = Thread(target=self.accept_request, args=(client, address))
            th.start()

    def accept_request(self, client_sock, client_addr):
        print(f"connect by {client_addr}")
        boNeedNoBreak = True
        while boNeedNoBreak:
            try:
                data = client_sock.recv(4096)
                req = data.decode("utf-8")
                print("-------req:" + req)
                response = "ack"
                #print("---response:\r\n" + response)
                client_sock.send(response.encode("utf-8"))
            except Exception as e:
                # clean up
                print("exception info: %s" %e)
                client_sock.shutdown(1)
                client_sock.close()
                boNeedNoBreak = False                      

