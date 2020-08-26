# Ringtail Export Converter
# By Andre Abadi

# file io imports
import os
import glob
import shutil
from datetime import datetime

# pandas imports
import numpy
import pandas

# import xlrd to catch pandas file read errors
import xlrd
from xlrd import XLRDError

# program version number, for use in MOTD upon execution
VERSION = 1.2


# set working directories
inputDir = "input/"
processedDir = "processed/"
outputDir = "output/"

# create the working folders (as above) in the executable directory
# /input
# /processed
# /output
def createFolders():
    if not os.path.exists(inputDir):
        os.makedirs(inputDir)
    if not os.path.exists(processedDir):
        os.makedirs(processedDir)
    if not os.path.exists(outputDir):
        os.makedirs(outputDir)


# transpose excel file into a neat data frame
def cleanColNames(temp):
    # drop the first five rows that always contain export metadata
    temp = temp.iloc[
        4:,
    ]
    # rename the column headings according to the first cell of each column
    counter = 0
    while counter < len(temp.columns):
        temp = temp.rename(columns={temp.columns[counter]: temp.iloc[0, counter]})
        # print(str(temp.iloc[0,counter]))
        counter = counter + 1
    # get rid of the first row now that their content is now column headings
    temp = temp.iloc[
        1:,
    ]
    # reset the row index after getting rid of the row just previously
    temp = temp.reset_index(drop=True)
    return temp



# prepend non-blank people with their people type
def prependType2(dataFrame):
    # https://stackoverflow.com/a/27275344
    people = [column for column in dataFrame if column.startswith("People/Organization")]
    for column in people:
        #print(dataFrame[column])
        heading = str(column)
        # split taking what's right of the space
        heading = heading.split(" ", 1)[-1]
        # add colon for formatting
        heading = heading + ": "
        dataFrame[column] = heading + dataFrame[column]
        # replace Nan with blanks
        dataFrame[column] = dataFrame[column].fillna("")
    return dataFrame

# updated people concatenation, does not depend on any ilocs
def concatPeople2(dataFrame):
    dataFrame["People"] = ""
    people = [column for column in dataFrame if column.startswith("People/Organization")]
    for column in people:
        dataFrame["People"] = dataFrame["People"] + " " + dataFrame[column]
    #print(dataFrame["People"])
    return dataFrame

# swap two columns by their names
# https://stackoverflow.com/a/56693510
def switchColumns(df, column1, column2):
    i = list(df.columns)
    a, b = i.index(column1), i.index(column2)
    i[b], i[a] = i[a], i[b]
    df = df[i]
    return df


# clean up the Document Date
def dateFormat(temp):
    dates = [column for column in temp if temp[column].str.contains("Date",regex=False)]
    for column in dates:
        # reformat datetime to a string in conventional format dd/MM/YYYY
        temp[column] = temp[column].apply(
            lambda x: x.strftime("%d/%m/%Y") if not pandas.isnull(x) else ""
        )
    return temp


# container function for all the smaller processing functions
def processData(temp):
    # get rid of the export metadata and assign proper column indexes
    temp = cleanColNames(temp)
    # if not blank prepend the people type to the cells
    temp = prependType2(temp)
    # combine all the people into one cell, ignoring blanks
    temp = concatPeople2(temp)
    # drop columns that start with "People/Organization" (old people columns)
    oldpeople = [column for column in temp if column.startswith("People/Organization")]
    for column in oldpeople:
        #print(column)
        temp = temp.drop(columns=column)
    # rename the "count" column
    temp = temp.rename(columns={"Count": "Item No."})
    # swap the people and document ID columns
    temp = switchColumns(temp, "People", "Document ID")
    # remove the time from the date/time stamp
    print(temp.columns)
    temp = dateFormat(temp)
    #temp["Document Date"] = pandas.to_datetime(temp["Document Date"])
    # add the page numbers column and leave it blank
    temp.insert(len(temp.columns), "Page Numbers", "")
    return temp


# add annexure column
def addAnnexure(temp):
    # get the affidavit prefix
    affidavitPrefix = ""
    affidavitPrefix = input("Enter annexure prefix (blank for no none): ")
    # convert to upper case
    affidavitPrefix = affidavitPrefix.upper()
    if affidavitPrefix == "":
        print("Annexure No. will be left blank")
        temp.insert(1, "Annexure No.", "")
    else:
        print("Annexure prefix will be: " + str(affidavitPrefix))
        # get the affidavit index
        affidavitIndex = 1
        affidavitIndex = input(
            "Enter annexure starting number (default is " + str(affidavitIndex) + "): "
        )
        # if blank recast string back to integer
        if affidavitIndex == "":
            affidavitIndex = 1
        print("Affidavit index will start at: " + str(affidavitIndex))
        affidavitIndex = int(affidavitIndex)
        # insert the affidavit ID as a new column
        temp.insert(
            1, "Annexure No.", range(affidavitIndex, affidavitIndex + len(temp))
        )
        # prefix the new column with the given prefix
        temp["Annexure No."] = affidavitPrefix + temp["Annexure No."].astype(str)
    return temp


# rename columns to make the names shorter
def shortenColNames(sheet):
    sheet.rename(columns={"Item No.": "Item"}, inplace=True)
    sheet.rename(columns={"Annexure No.": "Annex."}, inplace=True)
    sheet.rename(columns={"Document Date": "Doc Date"}, inplace=True)
    sheet.rename(columns={"Document Type": "Doc Type"}, inplace=True)
    sheet.rename(columns={"Document ID": "Doc ID"}, inplace=True)
    sheet.rename(columns={"Page Numbers": "Pages"}, inplace=True)
    return sheet


# output court book to file, setting column width
def writeCourtBook(filename, dataFrame):
    # create a writer object holding destination settings
    writer = pandas.ExcelWriter(filename, engine="xlsxwriter")
    dataFrame.to_excel(writer, startrow=0, sheet_name="Output", index=False)
    # select the workbook from the file object
    workbook = writer.book
    # select the worksheet from the workbook object
    worksheet = writer.sheets["Output"]
    # set all columns to vertically align to top and wrap
    formatting = workbook.add_format({"valign": "top"})
    formatting.set_text_wrap()
    worksheet.set_column("A:G", None, formatting)
    # manually set column widths
    worksheet.set_column(0, 0, 4.43)
    worksheet.set_column(1, 1, 5.57)
    worksheet.set_column(2, 2, 9.86)
    worksheet.set_column(3, 3, 10.14)
    worksheet.set_column(4, 4, 20)
    worksheet.set_column(5, 5, 20.14)
    worksheet.set_column(6, 6, 17.29)
    worksheet.set_column(7, 7, 4.86)
    # set the output to print/page view instead of "normal" view
    worksheet.set_page_view()
    # set margins to Excel "Narrow"
    mtop = 0.75  # inches
    msides = 0.25  # inches
    # convert margins to inches to comply with set_margins() function
    worksheet.set_margins(left=msides, right=msides, top=mtop, bottom=mtop)
    writer.save()

# wrapper function to create court book 
def createCourtBook(inputName, inputFile):
    # convert the ExcelFile object to a dataframe by using 1st (0th) sheet
    sheet = pandas.read_excel(inputFile, 0)
    # run all of the processing functions
    sheet = processData(sheet)
    # add annexure prefix (interactive)
    sheet = addAnnexure(sheet)
    # shorten the column names
    sheet = shortenColNames(sheet)
    # output to the designated file
    try:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        outputName = outputDir + timestamp + " - Court Book.xlsx"
        writeCourtBook(outputName, sheet)
        # sheet.to_excel(outputName)
        print("\nSuccessful output: " + outputName)
        # move the processed input file to the processed directory
        shutil.move(inputName, processedDir)
        # print("Moved '" + inputName + "' to '" + processedDir +"' folder")
    except FileNotFoundError:
        print("\nCannot write " + outputName)
    except PermissionError:
        print("\nFile Currently in use, please close it and try again.")
    print("")


def writeStatement(filename, dataFrame):
    # create a writer object holding destination settings
    writer = pandas.ExcelWriter(filename, engine="xlsxwriter")
    docID = dataFrame.filter(items=['Document ID'])
    # write just the ringtail references to the file
    docID.to_excel(writer, startrow=0, sheet_name="Output", index=False,header=False)
    # select the workbook from the file object
    workbook = writer.book
    # select the worksheet from the workbook object
    worksheet = writer.sheets["Output"]
    # set all columns to vertically align to top and wrap
    default = workbook.add_format({"valign": "top"})
    default.set_text_wrap()
    # create a new format name for the underline and bold format
    emphasis = workbook.add_format({'bold': True, 'underline': True,'valign':'top'})
    bold = workbook.add_format({'bold': True, 'valign':'top'})
    italic = workbook.add_format({'italic': True, 'valign':'top'})
    # emphasise the first column
    worksheet.set_column("A:A", None, emphasis)
    # creat a new dataframe with just the paragraph text
    dataFrame = dataFrame.filter(items=['Statement1','Statement2','Statement3','Statement4','Statement5','Statement6'])
    # iterate over each row
    # https://stackoverflow.com/a/16476974
    counter = 0
    for index, row in dataFrame.iterrows():
        worksheet.write_rich_string(counter,1,emphasis,row['Statement1'],row['Statement2'],emphasis,row['Statement3'],row['Statement4'],italic,row['Statement5'],row['Statement6'])
        counter = counter + 1
    # ensure that all cells get valign top and text wrap
    worksheet.set_column("A:B",None,default)
    # overwrite Document ID
    worksheet.set_column("A:A",None,bold)
    # set some column widths for readability
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 75)
    # finally save (not necessary but neatest)
    writer.save()


# wrapper function to create statement
def createStatement(inputName, inputFile):
    # convert the ExcelFile object to a dataframe by using 1st (0th) sheet
    sheet = pandas.read_excel(inputFile, 0)
    # clean up the excel file into a neat dataframe
    sheet = cleanColNames(sheet)
    # start off the new column
    sheet['Statement1'] = "I AM SHOWN"
    # with barcode
    sheet['Statement2'] = " a document barcoded "
    sheet['Statement2'] = sheet['Statement2'].astype(str) + sheet['Document ID']
    # lower-cased document type
    sheet['Statement2'] = sheet['Statement2'].astype(str) + " which "
    sheet['Statement3'] = "I IDENTIFY"
    sheet['Document Type'] = sheet["Document Type"].str.lower()
    sheet['Statement4'] = " as a " + sheet['Document Type']
    # title
    sheet['Statement4'] = sheet['Statement4'].astype(str) + " titled '"
    sheet['Statement5'] = sheet['Title']
    # if the date is under column name 'Document Date'
    dateName = 'Document Date'
    # otherwise it will be called 'Date
    if 'Date' in sheet:
        dateName = 'Date'
    # convert any strings to datetime values
    try:
        sheet[dateName] = pandas.to_datetime(sheet[dateName])
    except ValueError:
        # remove "Australian X Time"
        temp = sheet[dateName].str.split("Aus", n=1, expand=True)
        sheet[dateName] = temp[0]
        # try again to convert to datetime
        sheet[dateName] = pandas.to_datetime(sheet[dateName])
    # convert to long format date
    sheet[dateName] = sheet[dateName].dt.strftime("%#d %B %Y")
    # if date exists, prepend with 'and dated'
    sheet[dateName] = " and dated " + sheet[dateName].astype(str)
    # delete "and dated NaT" with a blank string
    sheet[dateName] = sheet[dateName].str.replace(" and dated NaT","",regex=False)
    # delete "and dated NaN" with a blank string
    sheet[dateName] = sheet[dateName].str.replace(" and dated nan","",regex=False)
    # then fill all empty date rows with blank strings
    sheet = sheet.fillna("")
    # append the date to the statement
    sheet['Statement6'] = "'"  + sheet[dateName]
    # add a full stop at the end no matter what
    sheet['Statement6'] = sheet['Statement6'].astype(str)  + "."
    try:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        outputName = outputDir + timestamp + " - Statement.xlsx"
        writeStatement(outputName, sheet)
        # sheet.to_excel(outputName)
        print("\nSuccessful output: " + outputName)
        # move the processed input file to the processed directory
        shutil.move(inputName, processedDir)
        # print("Moved '" + inputName + "' to '" + processedDir +"' folder")
    except FileNotFoundError:
        print("\nCannot write " + outputName)
    except PermissionError:
        print("\nFile Currently in use, please close it and try again.")
    print("")

# testing function that moves processed file back into input
def testReset():
    # if inputDir is empty
    if not os.listdir(inputDir):
        # move the first processed file back into input directory
        try:
            processedName = processedDir + os.listdir(processedDir)[0]
            shutil.move(processedName,inputDir)
            print("\nReset '" + processedName + "' to '" + inputDir + "'")
        except IndexError:
            print("\nNo files reset from \processed to \input")
        except PermissionError:
            print("\n" + processedName + "currently open. Please close.")
            quit()

# main function, wraps all the things
def main():
    try:
        print("Ringtail Export Converter by Andre Abadi, Version " + str(VERSION))
        # create file management folders if they don't already exist
        createFolders()
        # declare these variables so they have scope of all the main function
        inputName = None
        inputFile = None
        # TESTING ONLY - RESET PROCESSED FILE TO INPUT
        testReset()
        # open the file in pandas
        try:
            inputName = inputDir + os.listdir(inputDir)[0]
        except (FileNotFoundError, IndexError):
            print("\nCannot find any files in '" + inputDir + "'")
            quit()
        try:
            inputFile = pandas.ExcelFile(inputName)
        except xlrd.biffh.XLRDError:
            print("'" + inputName + "' is not a valid Excel file\n")
            quit()
        print("\nProcessing: " + inputName)
        choice = input("\n'1' for Court Book\n'2' for Statement\n(Default is Exit): ")
        if (choice == ""):
            print("Exiting\n")
            quit()
        if (choice == "1"):
            createCourtBook(inputName,inputFile)
        elif (choice == "2"):
            createStatement(inputName,inputFile)
    # catch CTRL+C and try to exit gracefully
    except KeyboardInterrupt:
        exit


# main function caller via implicit variable
if __name__ == "__main__":
    main()
