import mimetypes
import socket
import os
from sre_constants import SRE_FLAG_VERBOSE
import stat
from urllib.parse import unquote

from threading import Thread

NEWLINE = "\r\n"

def get_file_contents(file_name):
    #print("get_file_contents:" + file_name)
    with open(file_name, 'r', encoding='utf-8') as f:
        return f.read()

def get_file_binary_contents(file_name):
    with open(file_name, "rb") as f:
        return f.read()

def has_permission_other(file_name):
    stmode = os.stat(file_name).st_mode
    return getattr(stat, "S_IROTH") & stmode > 0

binary_type_files = set(["jpg", "jpeg", "mp3", "png", "html", "js", "css", "ico"])

def should_return_binary(file_extension):
    return file_extension in binary_type_files

mime_types = {
    "html": "text/html",
    "css": "text/css",
    "js": "text/javascript",
    "mp3": "audio/mpeg",
    "png": "image/png",
    "jpg": "image/jpg",
    "jpeg": "image/jpeg",
    "txt": "text/plain",
    "ico": "image/x-icon",
}

def get_file_mime_type(file_extension):
    mime_type = mime_types[file_extension]
    return mime_type if mime_type is not None else "text/plain"

class HTTPServer:
    """
    Our actual HTTP server which will service GET and POST requests.
    """

    def __init__(self, host="localhost", port=9001, directory="/"):
        print(f"Server started. Listening at http://{host}:{port}/")
        self.host = host
        self.port = port
        self.working_dir = os.path.dirname(os.path.realpath(__file__))
        self.working_dir += '\\';
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
            th = Thread(target=self.accept_request, args=(client, address))
            th.start()

    def accept_request(self, client_sock, client_addr):
        #print("-------accept_request")
        data = client_sock.recv(4096)
        req = data.decode("utf-8")
        #print("-------req:" + req)
        response = self.process_response(req)
        #print("---response:\r\n" + response)
        client_sock.send(response)

        # clean up
        client_sock.shutdown(1)
        client_sock.close()

    def process_response(self, request):
        formatted_data = request.strip().split(NEWLINE)
        request_words = formatted_data[0].split()

        if len(request_words) == 0:
            return

        requested_file = request_words[1][1:]
        if request_words[0] == "GET":
            return self.get_request(requested_file, formatted_data)
        if request_words[0] == "POST":
            return self.post_request(requested_file, formatted_data)
        return self.method_not_allowed()
    
     # TODO: Write the response to a GET request
    def get_request(self, requested_file, data):
        if(requested_file == ""):
            return self.resource_not_found()

        requested_file = self.working_dir + requested_file
        #print(requested_file)
        if (not os.path.exists(requested_file)):
            print("file not exists")
            return self.resource_not_found()
        elif (not has_permission_other(requested_file)):
            print("file no permission")
            return self.resource_forbidden()
        else:
            #print("start read file")
            builder = ResponseBuilder()

            if (should_return_binary(requested_file.split(".")[1])):
                #print("read binary")
                builder.set_content(get_file_binary_contents(requested_file))
            else:
                #print("read content")
                builder.set_content(get_file_contents(requested_file))
            
            builder.set_status("200", "OK")

            builder.add_header("Connection", "close")
            #print("----builder.add_header Connection close")
            builder.add_header("Content-Type", get_file_mime_type(requested_file.split(".")[1]))
            #print("----return builder.build()")
            return builder.build()
    # TODO: Write the response to a POST request
    def post_request(self, requested_file, data):
        
        builder = ResponseBuilder()
        builder.set_status("200", "OK")
        builder.add_header("Connection", "close")
        builder.add_header("Content-Type", mime_types["html"])
        print("---requested_file:" + requested_file)
        if(requested_file == "happybirthday"):
            for i in data:
                print("----data:" + i)
            
            builder.set_content("success")
        else:
            builder.set_content("unknown")
            
        return builder.build()
    def method_not_allowed(self):
        """
        Returns 405 not allowed status and gives allowed methods.
        TODO: If you are not going to complete the `ResponseBuilder`,
        This must be rewritten.
        """
        builder = ResponseBuilder()
        builder.set_status("405", "METHOD NOT ALLOWED")
        allowed = ", ".join(["GET", "POST"])
        builder.add_header("Allow", allowed)
        builder.add_header("Connection", "close")
        return builder.build()

    # TODO: Make a function that handles not found error
    def resource_not_found(self):
        """
        Returns 404 not found status and sends back our 404.html page.
        """
        builder = ResponseBuilder()
        builder.set_status("404", "NOT FOUND")
        builder.add_header("Connection", "close")
        builder.add_header("Content-Type", mime_types["html"])
        builder.set_content(get_file_contents("common/404.html"))
        return builder.build()

    # TODO: Make a function that handles forbidden error
    def resource_forbidden(self):
        """
        Returns 403 FORBIDDEN status and sends back our 403.html page.
        """
        builder = ResponseBuilder()
        builder.set_status("403", "FORBIDDEN")
        builder.add_header("Connection", "close")
        builder.add_header("Content-Type", mime_types["html"])
        builder.set_content(get_file_contents("403.html"))
        return builder.build()
    
class ResponseBuilder:
    def __init__(self):
        """
        Initialize the parts of a response to nothing.
        """
        self.headers = []
        self.status = None
        self.content = None

    def add_header(self, headerKey, headerValue):
        """ Adds a new header to the response """
        self.headers.append(f"{headerKey}: {headerValue}")

    def set_status(self, statusCode, statusMessage):
        """ Sets the status of the response """
        self.status = f"HTTP/1.1 {statusCode} {statusMessage}"

    def set_content(self, content):
        """ Sets `self.content` to the bytes of the content """
        #print("output content:")
        #print(content)
        if isinstance(content, (bytes, bytearray)):
            self.content = content
        else:
            self.content = content.encode("utf-8")

    # TODO Complete the build function
    def build(self):
       
        response = self.status
        response += NEWLINE
        for i in self.headers:
            response += i
        response += NEWLINE
        response += NEWLINE
        response = response.encode("utf-8")
        response += self.content
        
        return response
    

if __name__=="__main__":
    svr = HTTPServer()

    