from PyPDF2 import PdfFileMerger
import os,sys, getopt, comtypes.client

merger = PdfFileMerger()

def main(argv):
    try:
        opts, _ = getopt.getopt(argv, "hmci:o:", ["merge=", "convert=", "input=", "output="])
    except:
        print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
        sys.exit()

    if argv[0] == ("-m" or "--merge"):
        pdfmerger(opts)

    if argv[0] == ("-c" or "--convert"):
        converter(opts)

def pdfmerger(opts):
    inputfile = ''
    outputfile = ''

    for opt, arg in opts:
        if opt == '-h':
            print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
            sys.exit()

        elif opt in ("-i", "--input"):
            inputfile = arg
            for items in inputfile.split():
                if items.endswith('.pdf'):
                    merger.append(items)

        elif opt in ("-o", "--output"):
            outputfile = arg
            if outputfile.endswith('.pdf'):
                merger.write(outputfile)
            else:
                merger.write(f'{outputfile}.pdf')

    print('Done')
    sys.exit()

def converter(opts):
    inputfile = ''
    outputfile = ''
    wdFormatPDF = 17

    for opt, arg in opts:
        if opt == '-h':
            print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
            sys.exit()

        elif opt in ("-i", "--input"):
            inputfile = os.path.abspath(arg)
            word = comtypes.client.CreateObject('Word.Application')
            for items in inputfile.split():
                doc = word.Documents.Open(items)
                if opt in("-o", "--output"):
                    outputfile = arg
                    doc.SaveAs(outputfile, wdFormatPDF)
                    doc.Close()
                    word.Quit()
                else:
                    doc.SaveAs(inputfile.split('.')[0], wdFormatPDF)
                    doc.Close()
                    word.Quit()

if __name__ == "__main__":
    main(sys.argv[1:])
