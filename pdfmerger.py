from PyPDF2 import PdfFileMerger
import os,sys, getopt, comtypes.client

merger = PdfFileMerger()

def main(argv):
    try:
        opts, _ = getopt.getopt(argv, "hmci:o:", ["help=","merge=", "convert=", "input=", "output="])
    except:
        print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
        sys.exit()

    try:
        if argv[0] == ("-m" or "--merge"):
            pdfmerger(opts)

        elif argv[0] == ("-c" or "--convert"):
            converter(opts)

        elif argv[0] == ("-h" or "--help"):
            print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
            sys.exit()

        else:
            print("-h or --help for help")
            print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
            sys.exit()
    
    except IndexError:
        print("-h or --help for help")
        print("pdfmerger.py [-m | -c] -i <inputfiles> -o <outputfiles>")
        sys.exit()

def pdfmerger(opts):
    for opt, arg in opts:

        if opt in ("-i", "--input"):
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
        
        else:
            merger.write('merged.pdf')

    print('Done')
    sys.exit()

def converter(opts):
    wdFormatPDF = 17

    if '-i' in opts[1]:
        inputFile = os.path.abspath(opts[1][1])
        word = comtypes.client.CreateObject('Word.Application')
        for items in inputFile.split():
            doc = word.Documents.Open(items)
            try:
                outputFile = os.path.abspath(opts[2][1])
                doc.SaveAs(outputFile.split('.')[0], wdFormatPDF)
                doc.Close()
                word.Quit()
            except IndexError:
                doc.SaveAs(inputFile.split('.')[0], wdFormatPDF)
                doc.Close()
                word.Quit()
    print('Done')
    sys.exit()

if __name__ == "__main__":
    main(sys.argv[1:])
