import socket
import sys

HOST, PORT = "172.20.100.99", 5000
data = " ".join(sys.argv[1:])

# Create a socket (SOCK_STREAM means a TCP socket)
with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
    # Connect to server and send data
    sock.connect((HOST, PORT))
    sock.sendall(bytes(data + "한글\n", "utf-8"))

    # Receive data from the server and shut down
    received = str(sock.recv(1024), "utf-8")

# -*- coding: utf8 -*-
print("Sent:     {}".format(data))
print("Received: {}".format(received))