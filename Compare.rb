require 'spreadsheet'
require 'fileutils'

Spreadsheet.client_encoding = 'UTF-8'

book1 = Spreadsheet.open 'C:\Users\brackm1\Documents\Ruby Tools\Portfolio Page Comparator\IL_NAICOA_20170729_293.xls'
book2 = Spreadsheet.open 'C:\Users\brackm1\Documents\Ruby Tools\Portfolio Page Comparator\IL_NAICOA_20170729_293_2.xls'
book3 = Spreadsheet::Worksheet.new  

def compareSheet sheet1, sheet2, book3
	sheet1.rows.each_with_index do |row1, i|
		row2 = sheet2.row i
		compareRow(row1, row2)
	end
end

def compareRow row1, row2
	notMatchingFormat = Spreadsheet::Format.new(border: :thin, border_color: :red)
	row1.each_with_index do |cell1, i|
		cell2 = row2[i]
		unless cell1.to_s == cell2.to_s
			row2.set_format(i, notMatchingFormat)
		end
	end
end

book1.worksheets.each_with_index do |sheet1, i|
	sheet2 = book2.worksheet i
	compareSheet(sheet1, sheet2, book3)
end

book2.write 'C:\Users\brackm1\Documents\Ruby Tools\Portfolio Page Comparator\IL_NAICOA_20170729_293_compare.xls'
