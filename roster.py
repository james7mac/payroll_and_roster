import tabula
tabula.convert_into("MasterRoster.pdf", 'rostercsv.csv', output_format="csv", pages='all')