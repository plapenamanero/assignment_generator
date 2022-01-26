from PyQt5.QtWidgets import QFileDialog, QApplication


def csv_file(initial='./'):
    """ Opens a dialog to select a CSV file.

    Args:
        initial (str, optional): Initial path. Defaults to './'.

    Returns:
        str: Selected file path.
    """
    ext = 'csv'
    return gui_file(initial, ext)


def pdf_file(initial='./'):
    """ Opens a dialog to select a PDF file.

    Args:
        initial (str, optional): Initial path. Defaults to './'.

    Returns:
        str: Selected file path.
    """
    ext = 'pdf'
    return gui_file(initial, ext)


def excel_file(initial='./'):
    """ Opens a dialog to select a XLSX file.

    Args:
        initial (str, optional): Initial path. Defaults to './'.

    Returns:
        str: Selected file path.
    """
    ext = 'xlsx'
    return gui_file(initial, ext)


def gui_file(initial='./', ext='csv'):
    """ Opens a dialog to select a file using PyQt5

    Args:
        initial (str, optional): Initial path. Defaults to './'.
        ext (str, optional): File extension. Defaults to 'csv'.

    Returns:
        str: Selected file path.
    """
    app = QApplication([initial])
    fname = QFileDialog.getOpenFileName(None,
                                        "Select file",
                                        initial,
                                        filter=f"{ext} files (*.{ext})")

    return fname[0].strip().replace('\n', '')
