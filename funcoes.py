from spire.doc import *
from spire.doc.common import *
import shutil


def formataCpf(cpf) :
    if len(cpf) < 11:
        cpf = cpf.zfill(11)
    cpf = '{}.{}.{}-{}'.format(cpf[:3], cpf[3:6], cpf[6:9], cpf[9:])
    return cpf

def montarPdf(nomeDocx):
    print()
    pdf = Document()
    # Load a Word DOCX file
    pdf.LoadFromFile(nomeDocx)
    nomePdf = nomeDocx[: len(nomeDocx) - 4]
    nomePdf = nomePdf +"pdf"


    pdf.SaveToFile(f"{nomePdf}", FileFormat.PDF)
    pdf.Close()
    os.remove(nomeDocx)
    os.rename(nomePdf, 'contratos/'+nomePdf)
