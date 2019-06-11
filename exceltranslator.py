from cellwriter import Writer
import mtranslate
import xlrd

LANG = {"Turkce": "tr", "English": "en"}


class Translator():
    def __init__(self, window):
        self.restart = False
        self.window = window
        self.lang, self.fromRow, self.toRow, self.toColumn = self.window.getInputs()
        self.translatedXLS = []

        for xls in self.window.directory:

            excelFormat = xls.split(".")[1]

            # If selected file is a XLS
            if excelFormat == "xls":
                self.wrt = writer(xls)
                workbook = xlrd.open_workbook(xls)
                self.sheet = workbook.sheet_by_index(0)
                self.run(LANG[self.lang], xls.split("/")[-1])

                if self.restart:
                    break

                self.translatedXLS.append(xls.split("/")[-1])

            # TODO: SUPPORT
            elif excelFormat == "xlsx" or excelFormat == "xlsm":
                self.window.message("info", "{0} file not supporting.")
                self.restart = True
                self.window.restart()

        if not self.restart:
            print("Translation completed!")
            self.window.message("info", "Translation completed!")
            self.window.restart()
            self.restart = False

    def translate(self, lang, xls):

    def run(self, lang, xls):
        for i in range(self.strtSrc - 1, self.sheet.nrows):
            self.window.updateProgBar(i, self.sheet.nrows, len(self.translatedXLS), xls)
            try:
                translatedText = mtranslate.translate(self.sheet.cell_value(i, self.srcCol), lang, "auto")

            # If getting error in translating
            except:
                while True:
                    result = self.window.message("retry", "{0}\n{1}".format("Translation wasn't complete!",
                                                                            "Please check your internet connection and try again."))
                    if result:
                        try:
                            translatedText = mtranslate.translate(self.sheet.cell_value(i, self.srcCol), lang, "auto")
                            break
                        except:
                            pass
                    else:
                        self.restart = True
                        self.window.restart()
                        break
            if self.restart:
                break
            self.wrt.write(i, self.trgtCol, translatedText)
        self.wrt.save()
