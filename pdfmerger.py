from PyPDF2 import PdfFileMerger
import sys, getopt

merger = PdfFileMerger()

def main(argv):
    inputfile = ''
    outputfile = ''
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["ifile=", "ofile="])
    except:
        print("pdfmerger.py -i \"<inputfile_1> <inputfile_2> ... <inputfile_n>\" -o <outputfile>")
        sys.exit()

    for opt, arg in opts:
        if opt == '-h':
            print("pdfmerger.py -i \"<inputfile_1> <inputfile_2> ... <inputfile_n>\" -o <outputfile>")
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
            for items in inputfile.split():
                if items.endswith('.pdf'):
                    merger.append(items)

        elif opt in ("-o", "--ofile"):
            outputfile = arg
            if outputfile.endswith('.pdf'):
                merger.write(outputfile)
            else:
                merger.write(f'{outputfile}.pdf')

    print(f'Input files : {inputfile}\n- are merged into : {outputfile}')


if __name__ == "__main__":
    main(sys.argv[1:])
