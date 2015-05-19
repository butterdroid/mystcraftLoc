__author__ = 'Carlos'

import xlsxwriter
from os import listdir
from copy import deepcopy

# Create Workbook and Sheet
locDB = xlsxwriter.Workbook('locDB.xlsx')
sheet = locDB.add_worksheet()

# Generate Formats

headers = locDB.add_format({'bold': True,
                            'font_size': 12,
                            'bg_color': 'gray'})

noEnglish = locDB.add_format({'bg_color': 'purple'})

noTranslation = locDB.add_format({'bg_color': 'red'})


# Generate languages list out of files in folder, also generate language headers.
languageFiles = listdir('lang')
languages = []

for la in languageFiles:
    l = la.partition('.')[0]
    languages.append(l)
    # Headers in file
    sheet.write(0, languages.index(l) + 1, l, headers)

# Freeze language header row

sheet.freeze_panes(1, 0)

# loc dictionary to house all strings
loc = {}

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

# List with cells to format for missing English

missingEnglish = []


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
            print "stringID " + string[0] + " present in " + fileName + " missing in source."

            # try:
            #     loc[string[0]].update(s)
            # except KeyError:
            #     print "stringID "+string[0]+" present in "+fileName+" missing in source."
            # except:
            #     print "Weird!", sys.exc_info()[0]
            # else:
            #     loc[string[0]] = s

    readerLoc.close()

# Spit into Single File. FIGS currently,

row = 1

for stringID in loc:
    sheet.write(row, 0, stringID)

    # if 'en_US.lang' not in loc[stringID]:
    #    sheet.set_row(row, None, noEnglish)

    col = 1
    for lang in languageFiles:
        sheet.write(row, col, loc.get(stringID).get(lang))
        col += 1
    row += 1

# Use conditional formatting to highlight all empty cells

sheet.conditional_format(1, languageFiles.index('en_US.lang')+1, loc.__len__(), languageFiles.index('en_US.lang')+1,
                         {'type': 'blanks',
                          'format': noEnglish})

sheet.conditional_format(1, 1, loc.__len__(), languages.__len__(), {'type': 'blanks',
                                                                    'format': noTranslation})

sheet.set_column(0, loc.__len__(), 30)

locDB.close()
