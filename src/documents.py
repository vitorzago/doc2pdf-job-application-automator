import os
import json

from PyPDF2 import PdfFileReader

STD_DIR = os.path.join(os.getcwd(), "src",
                       "std_application")
STD_DIR_CURRICULUM_VITAE = os.path.join(STD_DIR, "1_curriculum_vitae")
STD_DIR_CERTIFICATES = os.path.join(STD_DIR, "2_certificates")
STD_FILE_CERTICATES_INDICES = os.path.join(STD_DIR_CERTIFICATES, "index.json")
STD_APPLICATION_PHOTO = os.path.join(STD_DIR, "max_mustermann.png")
class Document:
    def __init__(self):
        pass

class CurriculumVitae(Document):
    def __init__(self):
        pass

class MotivationLetter(Document):
    def __init__(self):
        pass

class Certificate:

    def __init__(self, pdf_file):
        self.filename = pdf_file
        self.dir = STD_DIR_CERTIFICATES
        self.filepath = os.path.join(self.dir, self.filename)
        self.indices = os.path.join(self.dir, "index.json")
        self.number_of_pages = self.get_number_of_pages()

    def get_number_of_pages(self):
        pdf = PdfFileReader(open(self.filepath, 'rb'))
        return pdf.getNumPages()

    def get_name(self, language):
        json_file = open(self.indices,)
        indices = json.load(json_file)
        files = [values for key, values in indices.items() if key == language][0]["database"]
        try:
            self.name = [parameters["name"] for parameters in files if parameters["filename"] == self.filename][0]
        except:
            print(self.filename)
            ValueError("Above file is not in the incides file.")
