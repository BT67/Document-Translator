from pptx import Presentation
from googletrans import Translator
from httpcore import SyncHTTPProxy
from base64 import b64encode
from docx import Document
from pptx.util import Pt

# Important, googletrans library must be version: googletrans==3.1.0a0 or higher

# def build_proxy_headers(username, password):
#     userpass = (username.encode("utf-8"), password.encode("utf-8"))
#     token = b64encode(b":".join(userpass))
#     return [(b"Proxy-Authorization", b"Basic " + token)]
#
#
# port = 8080
# proxy_url = (b"http", b"<proxy_host_ip>", port, b'')
# proxy_headers = build_proxy_headers("<username>", "<password>")
# proxy = {"https": SyncHTTPProxy(proxy_url=proxy_url, proxy_headers=proxy_headers)}

# translator = Translator(service_urls=['translate.googleapis.com'], proxies=proxy)

translator = Translator(service_urls=['translate.googleapis.com'])


def case001():
    translate_slideshow("test.pptx", "ja", "en")

def case002():
    translate_word_document("test.docx", "ja", "en")


def translate_spreadsheet(file, from_lang, to_lang):
    print(f"Translating spreadsheet: {file} from {from_lang} to {to_lang}")


def translate_word_document(file, from_lang, to_lang):
    print(f"Translating word document: {file} from {from_lang} to {to_lang}")
    document = Document(file)
    for paragraph in document.paragraphs:
        paragraph.text = translator.translate(paragraph.text, dest=to_lang).text
    filename = file[:file.rfind('.')] + "_" + to_lang + ".docx"
    document.save(filename)


def translate_slideshow(file, from_lang, to_lang):
    print(f"Translating slideshow: {file} from {from_lang} to {to_lang} \n")
    presentation = Presentation(file)
    for slide in presentation.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                if len(shape.text) > 1:
                    shape.text = translator.translate(shape.text, dest=to_lang).text
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
    filename = file[:file.rfind('.')] + "_" + to_lang + ".pptx"
    presentation.save(filename)


def main():
    # case001()
    case002()


if __name__ == "__main__":
    main()

# TODO:
# Add translation for tables within powerpoint slides
