from cellwriter import Writer
import mtranslate
import string
import xlrd

LANG = {"Turkce": "tr", "English": "en"}
alphabet = string.ascii_uppercase

class Cell():
    def __init__(self, arg):
        self.row=arg[0]
        self.column=arg[1]
        
class Translator():
    def __init__(self, window):
        self.window = window

        self.lang, self.fromCell, self.toCell, self.toColumn = self.window.getInputs()
        self.fromCell=Cell(self.convertToIndex((self.fromCell)))
        self.toCell=Cell(self.convertToIndex((self.toCell)))
        self.toColumn=self.convertToIndex((self.toColumn+"0"))[1]

        self.translatedXLS = []


        for xls in self.window.directory:

            excelFormat = xls.split(".")[1]

            # If selected file is a XLS
            if excelFormat == "xls":
                self.wrt = Writer(xls)
                workbook = xlrd.open_workbook(xls)
                self.sheet = workbook.sheet_by_index(0)

                self.translate(LANG[self.lang], xls.split("/")[-1])

                self.translatedXLS.append(xls.split("/")[-1])

            # TODO: SUPPORT
            elif excelFormat == "xlsx" or excelFormat == "xlsm":
                self.window.message("info", "{0} file not supporting.")
                self.window.restart()


        print("Translation completed!")
        self.window.message("info", "Translation completed!")
        self.window.restart()



    def convertToIndex(self,text):
        if text in ["END","end"]:
            return [-99]*2

        text=text.upper()
        row,column="",0
        
        firstChar=True
        for char in text[::-1]:
            if char not in alphabet:
                row+=char
            else:
                if firstChar:
                    column+=alphabet.index(char)
                    firstChar=False
                else :
                    column+=(alphabet.index(char)+1)*len(alphabet)


        return [int(row[::-1])-1,column]


    def translate(self, lang, xls):

        
        self.toCell.row = self.sheet.nrows if self.toCell.row == -99 else self.toCell.row+1

        for i in range(self.fromCell.row,self.toCell.row):
            self.window.updateProgBar(i,self.fromCell.row, self.toCell.row, len(self.translatedXLS), xls)

            try:
                translatedText = mtranslate.translate(self.sheet.cell_value(i, self.fromCell.column), lang, "auto")

            # If getting error in translating
            except Exception as e:
                while True:
                    print(e)
                    result = self.window.message("retry", "{0}\n{1}".format("Translation wasn't complete!",
                                                                            "Please check your internet connection and try again."))
                    if result:
                        try:
                            translatedText = mtranslate.translate(self.sheet.cell_value(i, self.fromCell.column), lang, "auto")
                            break
                        except:
                            pass
                    else:
                        self.window.restart()
                        self.wrt.save()
                        return False    

            self.wrt.write(i, self.toColumn, translatedText)
        self.wrt.save()
        return True
