import socket
import openpyxl

workbook = openpyxl.load_workbook('\Users\Public\Documents\ENDEREÇOS.xlsx')
sheet = workbook.active


# entrando com os dados
var = '320800A7'

#Layot da etiquta

codigo_zpl = 'CT~~CD,~CC^~CT~  \n' \
            '^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4~SD10^JUS^LRN^CI0^XZ \n' \
            '^XA \n' \
            '^MMT \n' \
            '^PW799 \n' \
            '^LL0400 \n' \
            '^LS0 \n' \
            '^BY8,3,140^FT751,118^BCI,,N,N \n' \
            '^FD>:' + var + '^FS \n' \
            '^FT797,283^A0I,84,604^FH\^FD' + var + '^FS \n' \
            '^PQ1,0,1,Y^XZ \n' \



# Endereço IP da impressora
HOST = '192.168.107.128'
# Porta padrão de comunicação com a impressora
PORT = 9100

# Conecta-se à impressora
s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.connect((HOST, PORT))

# Envia o código ZPL para a impressora
s.sendall(codigo_zpl.encode())

# Fecha a conexão com a impressora
s.close()
