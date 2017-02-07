require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

book1 = Spreadsheet.open 'C:\Users\brackm1\Documents\Ruby Tools\Portfolio Page Comparator\IL_NAICOA_20170729_293.xls'
book2 = Spreadsheet.open 'C:\Users\brackm1\Documents\Ruby Tools\Portfolio Page Comparator\IL_NAICOA_20170729_293_2.xls' 

def compareSheet sheet1, sheet2
	@rowCount = 1
	sheet1.rows.each_with_index do |row1, i|
		row2 = sheet2.row i
		compareRow(row1, row2, sheet2)
		@rowCount += 1
	end
end



def compareRow row1, row2, sheet2
	notMatchingFormat = Spreadsheet::Format.new(border: :thin, border_color: :red)
	row1.each_with_index do |cell1, i|
		cell2 = row2[i]
		if cell1.to_s != cell2.to_s
			puts "#{sheet2.name}   ---   Column #{i + 1}   ---   Row #{@rowCount}   -->   #{row2[i]}"
		end
	end
end

book1.worksheets.each_with_index do |sheet1, i|
	s2 = book2.worksheet i
	sheet2 = s2.clone
	compareSheet(sheet1, sheet2)
end

#book2.write 'C:\Users\brackm1\Documents\Ruby Tools\Portfolio Page Comparator\IL_NAICOA_20170729_293_compare.xls'
