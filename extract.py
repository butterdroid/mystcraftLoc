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

# Generate languages list out of files in folder, also generate language headers
languageFiles = listdir('lang')
languages = []

sheet.write('A1', 'String Number', headers)
sheet.write('B1', 'String ID', headers)


for la in languageFiles:
    l = la.partition('.')[0]
    languages.append(l)
    # Headers in file
    sheet.write(0, languages.index(l) + 2, l, headers)

# Freeze language header row

sheet.freeze_panes(1, 0)

# loc dictionary to house all strings
loc = {}

# Pull English into loc (first, to find discrepancies)
readerUS = open("lang\en_US.lang", 'r')

stringNum = 1

for line in readerUS:
    string = line.partition('=')
    loc[string[0]] = {'num': stringNum, 'en_US.lang': string[2]}
    if len(line) == 1:
        loc['Empty-'+str(stringNum)] = {'num': stringNum}
    stringNum += 1


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
        if string[0] in loc:
            loc[string[0]].update({fileName: string[2].decode('utf-8').strip('\n')})
        else:
            loc[string[0]] = {'num': stringNum, fileName: string[2].decode('utf-8').strip('\n')}
            stringNum += 1
            print "stringID " + string[0] + " present in " + fileName + " missing in source."

    readerLoc.close()

# Spit into Single File. FIGS currently,



for stringID in loc:
    sheet.write(loc[stringID].get('num'), 1, stringID)
    sheet.write(loc[stringID].get('num'), 0, loc[stringID].get('num'))
    col = 2
    # Check if row is empty and hide it from view in spreadsheet.
    if stringID.partition('-')[0] == 'Empty':
        sheet.set_row(loc[stringID].get('num'), None, None, {'hidden': True})
    else:
        for lang in languageFiles:
            sheet.write(loc[stringID].get('num'), col, loc.get(stringID).get(lang))
            col += 1

# Use conditional formatting to highlight all empty cells

sheet.conditional_format(1, languageFiles.index('en_US.lang') + 2, loc.__len__(), languageFiles.index('en_US.lang') + 2,
                         {'type': 'blanks',
                          'format': noEnglish})

sheet.conditional_format(1, 1, loc.__len__(), languages.__len__(), {'type': 'blanks',
                                                                    'format': noTranslation})

# Hide ID column and stretch all columns to make it look cleaner

sheet.set_column(1, loc.__len__(), 30)

sheet.set_column('A:A', None, None, {'hidden': True})


locDB.close()
