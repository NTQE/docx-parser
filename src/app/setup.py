import os
from src.app.models.esipdata import EsipData
from src.app.parsers.default import docx_parser, txt_parser
from src.app.inserters.default import inserter
from src.app.manager import DocManager, TemplateManager
from src.app import config


def find_files(path: str, ext: str = '.docx') -> list[str]:
    """find files in directory with extension 'ext'. No recursion.

    :param path: folder to find files in
    :param ext: extension being searched for. requires '.' in the string
    :return: list of absolute file paths with extension 'ext' in folder 'path'
    """
    return [os.path.join(path, file) for file in os.listdir(path) if file.lower().endswith(ext)]


def gather_data(path: str, parse_docx, parse_text) -> EsipData:
    """Gathers data from single file using the parsing functions passed as arguments

    :param path: absolute file path to a .docx document
    :param parse_docx: docx parsing function
    :param parse_text: txt parsing function
    :return: an object containing data extracted with the parsers as an EsipData object
    """
    with (DocManager(path) as doc, open(f"{os.path.splitext(path)[0]}{'.txt'}") as txt):
        return parse_text(txt, parse_docx(doc, EsipData(abs_path=path)))


def gather_from_files(files: list[str]) -> list:
    """Handles gathering data from a list of absolute file paths using the gather_data function

    :param files: list of absolute file paths
    :return: list of EsipData objects
    """
    files_list = []
    for file in files:
        files_list.append(gather_data(file, docx_parser, txt_parser))
    return files_list


def insert_data(file_data, template_path, new_save_path):
    """ insert data into a blank template prepared ahead of time

    :param file_data: data from individual file gathered into EsipData objects
    :param template_path: path to the new empty template prepared ahead of time
    :param new_save_path: new directory to save the file to differentiate from the old
    :return: nothing. check the new_save_path folder for updated docs.
    """
    with TemplateManager(template_path, new_save_path) as template_doc:
        inserter(template_doc, file_data)


def insert_into_files(file_data_list):
    """handles the insertion of data into prepared template from list of file data

    :param file_data_list: list of data from files in the form of EsipData objects
    :return: nothing, check the new_save_path folder for updated docs.
    """
    base_path, template_path, new_path = config.configuration()
    for file_data in file_data_list:
        new_save_path = os.path.join(new_path, file_data.file_name)
        insert_data(file_data, template_path, new_save_path)