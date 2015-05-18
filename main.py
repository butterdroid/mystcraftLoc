__author__ = 'Carlos'

import xlsxwriter
from os import listdir
from copy import deepcopy

locDB = xlsxwriter.Workbook('locDB.xlsx')
sheet = locDB.add_worksheet()

loc = {}

# Generate languages list out of files in folder, also generate language headers.
languageFiles = listdir('lang')
languages = []

for la in languageFiles:
    languages.append(la.partition('.')[0])

for la in languages:
    sheet.write(0, languages.index(la) + 1, la)

# Pull English into loc (first, to find discrepancies)

readerUS = open("lang\en_US.lang", 'r')

for line in readerUS:
    string = line.partition('=')
    loc[string[0]] = {'en_US.lang': string[2]}

readerUS.close()

# Pull every other language into Loc. If stringID not already in loc, add it and create log.

# new lang list without US
langmE = deepcopy(languageFiles)
langmE.remove('en_US.lang')

# now actual function

for fileName in languageFiles:
    readerLoc = open("lang\\" + fileName, 'r')

    for line in readerLoc:
        string = line.partition('=')
        s = {fileName: string[2].decode('utf-8').strip('\n')}
        if string[0] in loc:
            loc[string[0]].update(s)
        else:
            loc[string[0]] = s
            print "stringID "+string[0]+" present in "+fileName+" missing in source."

        # try:
        #     loc[string[0]].update(s)
        # except KeyError:
        #     print "stringID "+string[0]+" present in "+fileName+" missing in source."
        # except:
        #     print "Weird!", sys.exc_info()[0]
        # else:
        #     loc[string[0]] = s

    readerLoc.close()

print loc

# Spit into Single File. FIGS currently,

row = 1

for stringID in loc:
    sheet.write(row, 0, stringID)
    col = 1
    for lang in languageFiles:
        sheet.write(row, col, loc.get(stringID).get(lang))
        col += 1
    row += 1

locDB.close()
