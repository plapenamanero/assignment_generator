from PyPDF2 import PdfFileWriter, PdfFileReader
from IPython.display import display
import ipywidgets as widgets
import threading
import os
import shutil
from . import gui


def create_pdfs(assignment):
    """ Creates individual PDF files from a file with all sheets.
        Ask for the original file using a file dialog

    Args:
        assignment (Assignment): Assignment object.
    """

    # number of pages per sheet
    n = int(assignment.config['Value'][7])

    # password to encryp the file
    password = assignment.config['Value'][8]

    # removes sheet folder contents
    for sheet in os.listdir('sheets'):
        sheet_file = os.path.join('sheets', sheet)
        try:
            if os.path.isfile(sheet_file) or os.path.islink(sheet_file):
                os.unlink(sheet_file)
            elif os.path.isdir(sheet_file):
                shutil.rmtree(sheet_file)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (sheet_file, e))

    # sets the password to False if it is empty
    if password.isspace() or (not password):
        password = False

    # opens dialog to ask for original file
    pdf_file = gui.pdf_file()

    print('------')
    print("Creating files")

    # creates pdf and show progress bar in Jupyter
    progress = widgets.FloatProgress(value=0.0, min=0.0, max=1.0)
    thread = threading.Thread(target=split_pdf,
                              args=(assignment,
                                    pdf_file, n, password, progress))
    display(progress)
    thread.start()


def split_pdf(assignment, pdf_file, n, password, progress):
    """ Splits pdf in multiple files giving the number of pages per document.

    Args:
        assignment (Assignment): Assignment object.
        pdf_file (str): Original pdf file.
        n (int): Number of pages per document.
        password (str or bool): Password to encrypt documents. If set to false
                                documents are no encrypted.
        progress (widget): Progress bar ipywidget

    Returns:
        [bool]: Returns True if the execution is successful.
    """
    data = assignment.student_list

    pdf = PdfFileReader(pdf_file)
    total = len(data)

    # creates individual documents
    for i in range(total):
        pdf_writer = PdfFileWriter()
        output_file = data['file'][i]

        for j in range(n):
            pdf_writer.addPage(pdf.getPage(n * i + j))

        if password:
            pdf_writer.encrypt(user_pwd=password,
                               owner_pwd=None,
                               use_128bit=True)

        with open(output_file, 'wb') as out:
            pdf_writer.write(out)

        progress.value = float(i + 1) / total

    print("Files created")

    return True
