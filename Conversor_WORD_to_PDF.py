#Conversor_WORD_to_PDF

import sys
import os
import comtypes.client

# Codigo correspondente ao formato .pdf
wdFormatPDF = 17

# Recupera o path completo (absoluto) do arquivo de
# entrada (.doc) a partir do primeiro parametro da
# linha de comando
in_file = os.path.abspath(sys.argv[1])

# Recupera o path completo (absoluto) do arquivo de
# saida (.pdf) a partir do segundo parametro da linha
# de comando
out_file = os.path.abspath(sys.argv[2])

# Cria instancia de um objeto COM para manipular Documentos Word
word = comtypes.client.CreateObject('Word.Application')

# Carrega Arquivo de entrada (.doc)
doc = word.Documents.Open(in_file)

# Salva arquivo de saida em formato .pdf
doc.SaveAs(out_file, FileFormat=wdFormatPDF)

# Fecha arquivo de Entrada
doc.Close()

# Finaliza instancia do Objeto COM criado
word.Quit()


#$ python doc2pdf.py entrada.doc saida.pdf