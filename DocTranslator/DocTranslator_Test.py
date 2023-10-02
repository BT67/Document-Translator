from pptx import Presentation
from googletrans import Translator

translator = Translator(service_urls=['translate.googleapis.com'])


def case001():
    translate_slideshow("test.pptx", "ja", "en")


def translate_spreadsheet(file, from_lang, to_lang):
    print(f"Translating spreadsheet: {file} from {from_lang} to {to_lang}")


def translate_document(file, from_lang, to_lang):
    print(f"Translating document: {file} from {from_lang} to {to_lang}")


def translate_slideshow(file, from_lang, to_lang):
    print(f"Translating slideshow: {file} from {from_lang} to {to_lang} \n")
    for slide in Presentation(file).slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                if len(shape.text) > 1:
                    print("Original text=" + shape.text)
                    print("Translated text=" + str(translator.translate(shape.text, dest=to_lang).text) + "\n")


def main():
    case001()


if __name__ == "__main__":
    main()
