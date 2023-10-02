import argparse
from pptx import Presentation
from googletrans import Translator

translator = Translator(service_urls=['translate.googleapis.com'])


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
    parser = argparse.ArgumentParser(description='Translate Documents')
    parser.add_argument('-i', '--input', help='Document to be translated')
    parser.add_argument('-f', '--from_lang', help='Document original language to be translated from (default Japanese)')
    parser.add_argument('-t', '--to_lang', help='Document target language to be translated to (default English)')
    args = parser.parse_args()

    if args.input is None:
        parser.print_help(file=None)
        return
    if args.from_lang is None:
        from_lang = 'ja'
    else:
        from_lang = args.from_lang
    if args.to_lang is None:
        to_lang = 'en'
    else:
        to_lang = args.to_lang

    file_ext = args.input[args.input.rfind('.'):]
    print(file_ext)
    match file_ext:
        case ".txt":
            translate_document(args.input, from_lang, to_lang)
        case ".xlsx":
            translate_spreadsheet(args.input, from_lang, to_lang)
        case ".pptx":
            translate_slideshow(args.input, from_lang, to_lang)


if __name__ == "__main__":
    main()
