require 'yaml'
require 'spreadsheet'

SETTING = YAML.load_file("settings.yml")
EXCELFILEPATH = SETTING["excelFilePath"]
COLUMNROW = SETTING["columnRow"]
TARGETSHEET = SETTING["targetSheet"]
TARGETROW = SETTING["targetRow"]
TEXTFILE = SETTING["textfile"]
OUTPUTMARGEPATH = SETTING["outputMargePath"]
OUTPUTMARGESHEETNAME = SETTING["outputMargeSheetName"]

Spreadsheet.client_encoding = 'UTF-8'

def makeExcel(excelHash, txtHash)
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet(:name => OUTPUTMARGESHEETNAME)

    i = 0
    excelHash.each_key do |key|
        sheet[0, i] = key
        sheet[1, i] = excelHash.fetch(key)
        sheet[2, i] = txtHash.fetch(key)
        i += 1
    end

    book.write(OUTPUTMARGEPATH)
end

def selectExcelData
    columnArray = []
    cellArray = []
    excelHash = {}

    book = Spreadsheet.open(EXCELFILEPATH, 'rb')

    book.worksheets.each do |ws|
        next if ws.name != TARGETSHEET
        ws.each_with_index do |row, row_idx|
            row.each do |cell|
                if row_idx == COLUMNROW - 1 then
                    columnArray.push cell
                elsif row_idx == TARGETROW - 1 then
                    cellArray.push cell
                end
            end
        end
        columnArray.each_with_index do |columnRow, idx|
            cell = cellArray[idx]
            excelHash.store(columnRow, cell) if cell != nil
        end
    end
    return excelHash
end

def selectTxtData
    txtHash = {}

    File.open(TEXTFILE).each do |row|
        commaArray = row.split(",")

        commaArray.each do |comma|
            equalArray = comma.split("=")
            txtHash.store(equalArray[0], equalArray[1])
        end
    end
    return txtHash
end

begin
    makeExcel(selectExcelData, selectTxtData)
rescue => ex
    puts "Error " + ex.message
    puts ex.backtrace
    exit
end