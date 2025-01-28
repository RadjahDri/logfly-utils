#!/usr/bin/env python3

import argparse
import datetime
import sqlite3
import xlsxwriter

from xlsxwriter.utility import xl_rowcol_to_cell

TIME_CATEGORIES_STEP = 60 * 60

GLIDER_SHEET_NAME = "Glider"
DURATION_SHEET_NAME = "Duration"
SITE_SHEET_NAME = "Site"
COUNTRY_SHEET_NAME = "Country"
TAG_SHEET_NAME = "Tag"

class Fly:
	def __init__(self, rawData):
		self.year = rawData[0][0:4]
		self.takeoffHour = rawData[0][8:]
		self.duration = rawData[1]
		self.gliderName = rawData[4]
		self.siteName = rawData[2]
		self.countryName = rawData[3]
		self.tags = convertCommentToTags(rawData[5])


def convertCommentToTags(comment):
    if(not comment):
        return []
    return list(filter(lambda x: x != "", map(lambda x: x.split(" ")[0], comment.split("\n"))))



def secondToTimeString(second):
	return "%dh%d" % (second//60//60, (second//60)%60)


def getFliesYears(dbPath):
	db = sqlite3.connect(dbPath)
	cur = db.cursor()
	reqResult = db.execute("SELECT DISTINCT SUBSTR(V_Date, 0, 5) FROM Vol;")
	years = list(map(lambda x: int(x[0]), reqResult.fetchall()))
	db.close()
	return years


def extractFlies(dbPath):
	result = {}
	years = getFliesYears(dbPath)

	db = sqlite3.connect(dbPath)
	cur = db.cursor()

	for year in years:
		reqResult = db.execute("SELECT V_Date, V_Duree, V_Site, S_Pays, V_Engin, V_Commentaire FROM Vol INNER JOIN Site ON V_Site = S_Nom WHERE '%d' = SUBSTR(V_Date, 0, 5);" % year)
		result[year] = list(map(lambda x: Fly(x), reqResult.fetchall()))

	db.close()

	return result


def classifyGeneric(flies, classifyFunc):
	categorisedFlies = {}

	for fly in flies:
		flyCategory = classifyFunc(fly)
		if(not flyCategory in categorisedFlies):
			categorisedFlies[flyCategory] = {'Times':0, "Duration":0}
		categorisedFlies[flyCategory]['Times'] += 1
		categorisedFlies[flyCategory]['Duration'] += fly.duration
	return categorisedFlies


def classifyArrayGeneric(flies, classifyFunc):
	categorisedFlies = {}
	
	categorisedFlies["Total"] = {'Times':0, "Duration":0}

	for fly in flies:
		flyCategories = classifyFunc(fly)
		for flyCategory in flyCategories:
			if(not flyCategory in categorisedFlies):
				categorisedFlies[flyCategory] = {'Times':0, "Duration":0}
			categorisedFlies[flyCategory]['Times'] += 1
			categorisedFlies[flyCategory]['Duration'] += fly.duration
		categorisedFlies["Total"]['Times'] += 1
		categorisedFlies["Total"]['Duration'] += fly.duration
	return categorisedFlies


def classifyByDurationCategories(flies):
	return classifyGeneric(flies, lambda fly: fly.duration // TIME_CATEGORIES_STEP)


def classifyByGlider(flies):
	return classifyGeneric(flies, lambda fly: fly.gliderName)


def classifyBySite(flies):
	return classifyGeneric(flies, lambda fly: fly.siteName)


def classifyByCountry(flies):
	return classifyGeneric(flies, lambda fly: fly.countryName)


def classifyByTags(flies):
	return classifyArrayGeneric(flies, lambda fly: fly.tags)


def exportGenericToXls(categorisedFlies, xlsFile, sheetName):
	sheet = xlsFile.add_worksheet(sheetName)
	chartSheet = xlsFile.add_worksheet("%sChart" % sheetName)
	chartIdx = 0
	rowIdx = 0
	for year in sorted(categorisedFlies.keys()):
		chart = xlsFile.add_chart({
			'type': 'column'
		})

		sheet.write_row(rowIdx, 0, (year,))
		rowIdx += 1
		titles = ["", "Nombre", "Temps"]
		sheet.write_row(rowIdx, 0, titles)
		beginRow = rowIdx
		rowIdx += 1

		for category, values in sorted(categorisedFlies[year].items()):
			sheet.write_row(rowIdx, 0, (category, values["Times"], values["Duration"]//60//60))
			rowIdx += 1

		endRow = rowIdx - 1

		chart.set_title({"name": str(year)})
		chart.add_series({
			'name': '=%s!%s' % (sheetName, xl_rowcol_to_cell(beginRow, 1, row_abs=True, col_abs=True)),
			'categories': '=%s!%s:%s' % (
				sheetName,
				xl_rowcol_to_cell(beginRow+1, 0, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 0, row_abs=True, col_abs=True)
			),
			'line':   {'none': True},
			'marker': {'type': 'automatic'},
			'values': '=%s!%s:%s' % (
				sheetName,
				xl_rowcol_to_cell(beginRow+1, 1, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 1, row_abs=True, col_abs=True)
			),
			'data_labels': {'value': True},
		})
		chart.add_series({
			'name': '=%s!%s' % (sheetName, xl_rowcol_to_cell(beginRow, 2, row_abs=True, col_abs=True)),
			'categories': '=%s!%s:%s' % (
				sheetName,
				xl_rowcol_to_cell(beginRow+1, 0, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 0, row_abs=True, col_abs=True)
			),
			'values': '=%s!%s:%s' % (
				sheetName,
				xl_rowcol_to_cell(beginRow+1, 2, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 2, row_abs=True, col_abs=True)
			),
			'data_labels': {'value': True},
		})

		sheet.autofit()

		chart.set_size({'x_scale': 2.5, 'y_scale': 1.5})

		chartSheet.insert_chart(xl_rowcol_to_cell(22 * chartIdx, 0), chart)
		chartIdx += 1


def exportDurationToXls(categorisedFlies, xlsFile):
	sheet = xlsFile.add_worksheet(DURATION_SHEET_NAME)
	chartSheet = xlsFile.add_worksheet("%sChart" % DURATION_SHEET_NAME)
	chartIdx = 0
	rowIdx = 0
	for year in sorted(categorisedFlies.keys()):
		chart = xlsFile.add_chart({
			'type': 'line'
		})

		sheet.write_row(rowIdx, 0, (year,))
		rowIdx += 1
		titles = ["Heure", "Nombre", "Temps"]
		sheet.write_row(rowIdx, 0, titles)
		beginRow = rowIdx
		rowIdx += 1

		for y in range(max(categorisedFlies[year].keys())):
			sheet.write_row(rowIdx+y, 0, ("%dh" % y, 0))

		for category, values in  categorisedFlies[year].items():
			sheet.write_row(rowIdx+category, 0, ("%dh" % category, values["Times"], values["Duration"]//60//60))

		rowIdx += max(categorisedFlies[year].keys())+1

		endRow = rowIdx - 1
		chart.set_title({"name": str(year)})
		chart.add_series({
			'name': '=%s!%s' % (DURATION_SHEET_NAME, xl_rowcol_to_cell(beginRow, 1, row_abs=True, col_abs=True)),
			'categories': '=%s!%s:%s' % (
				DURATION_SHEET_NAME,
				xl_rowcol_to_cell(beginRow+1, 0, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 0, row_abs=True, col_abs=True)
			),
			'values': '=%s!%s:%s' % (
				DURATION_SHEET_NAME,
				xl_rowcol_to_cell(beginRow+1, 1, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 1, row_abs=True, col_abs=True)
			),
			'data_labels': {'value': True},
		})
		chart.add_series({
			'name': '=%s!%s' % (DURATION_SHEET_NAME, xl_rowcol_to_cell(beginRow, 2, row_abs=True, col_abs=True)),
			'categories': '=%s!%s:%s' % (
				DURATION_SHEET_NAME,
				xl_rowcol_to_cell(beginRow+1, 0, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 0, row_abs=True, col_abs=True)
			),
			'values': '=%s!%s:%s' % (
				DURATION_SHEET_NAME,
				xl_rowcol_to_cell(beginRow+1, 2, row_abs=True, col_abs=True),
				xl_rowcol_to_cell(endRow, 2, row_abs=True, col_abs=True)
			),
			'data_labels': {'value': True},
			'y2_axis': True,
		})

		sheet.autofit()

		chartSheet.insert_chart(xl_rowcol_to_cell(15 * (chartIdx // 2), 8 * (chartIdx % 2)), chart)
		chartIdx += 1


def exportGliderToXls(categorisedFlies, xlsFile):
	exportGenericToXls(categorisedFlies, xlsFile, GLIDER_SHEET_NAME)


def exportSiteToXls(categorisedFlies, xlsFile):
	exportGenericToXls(categorisedFlies, xlsFile, SITE_SHEET_NAME)


def exportCountryToXls(categorisedFlies, xlsFile):
	exportGenericToXls(categorisedFlies, xlsFile, COUNTRY_SHEET_NAME)


def exportTagToXls(categorisedFlies, xlsFile):
	exportGenericToXls(categorisedFlies, xlsFile, TAG_SHEET_NAME)


def argumentParsing():
	parser = argparse.ArgumentParser()

	parser.add_argument("-d", "--db",
		type=str,
		required=True,
		help="Path to flies database")

	parser.add_argument("-o", "--output",
		type=str,
		required=True,
		help="Path to XLS output file")

	return parser.parse_args()


def main(args):
	flies = extractFlies(args.db)

	durationCategorisedFlies = {}
	gliderCategorisedFlies = {}
	siteCategorisedFlies = {}
	countryCategorisedFlies = {}
	tagsFlies = {}

	for year in flies.keys():
		print("=== Country ===")
		countryCategorisedFlies[year] = classifyByCountry(flies[year])
		print("= %d =" % year)
		for category, value in  countryCategorisedFlies[year].items():
			print("%s\t%d\t%s" % (category, value["Times"], secondToTimeString(value["Duration"])))

		print("=== Site ===")
		siteCategorisedFlies[year] = classifyBySite(flies[year])
		print("= %d =" % year)
		for category, value in  siteCategorisedFlies[year].items():
			print("%s\t%d\t%s" % (category, value["Times"], secondToTimeString(value["Duration"])))

		print("=== Glider ===")
		gliderCategorisedFlies[year] = classifyByGlider(flies[year])
		print("= %d =" % year)
		for category, value in  gliderCategorisedFlies[year].items():
			print("%s\t%d\t%s" % (category, value["Times"], secondToTimeString(value["Duration"])))

		print("=== Duration ===")
		durationCategorisedFlies[year] = classifyByDurationCategories(flies[year])
		print("= %d =" % year)
		for category, value in  durationCategorisedFlies[year].items():
			print("%dh\t%d\t%s" % (category, value["Times"], secondToTimeString(value["Duration"])))

		print("=== Tag ===")
		tagsFlies[year] = classifyByTags(flies[year])
		print("[*] %d" % len(tagsFlies))
		print("= %d =" % year)
		for category, value in tagsFlies[year].items():
			print("%s\t%d\t%s" % (category, value["Times"], secondToTimeString(value["Duration"])))


	with xlsxwriter.Workbook(args.output) as xlsFile:
		exportDurationToXls(durationCategorisedFlies, xlsFile)
		exportGliderToXls(gliderCategorisedFlies, xlsFile)
		exportSiteToXls(siteCategorisedFlies, xlsFile)
		exportCountryToXls(countryCategorisedFlies, xlsFile)
		exportTagToXls(tagsFlies, xlsFile)


if(__name__ == "__main__"):
	args = argumentParsing()
	exit(main(args))
