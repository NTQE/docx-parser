from docx import Document


class DocManager:
    """Context manager for gathering data using python-docx objects

    """
    def __init__(self, path: str):
        # print('Creating document object with python-docx')
        self.doc = Document(path)

    def __enter__(self):
        # print('Returning python-docx document object')
        return self.doc

    def __exit__(self, exc_type, exc_val, exc_tb):
        # print('Deleting python-docx document object')
        del self.doc


class TemplateManager:
    """Context manager for updating a template with gathered data using python-docx

    """
    def __init__(self, template_path: str, save_path: str):
        # print('Creating document object with python-docx')
        self.doc = Document(template_path)
        self.save_path = save_path

    def __enter__(self):
        # print('Returning python-docx document object')
        return self.doc

    def __exit__(self, exc_type, exc_val, exc_tb):
        # print('Deleting python-docx document object')
        self.doc.save(self.save_path)
        del self.doc
