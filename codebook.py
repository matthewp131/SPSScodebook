from savReaderWriter import SavHeaderReader
import pandas
import argparse
import os


def main():
    savFile, xlsxFile = get_args()
    metadata, report = get_report(savFile)
    frames = {}

    # Generate Variable Information dataframe
    df = pandas.DataFrame()
    df["Variable"] = metadata.varNames
    df["Position"] = list(range(1, df.shape[0] + 1))
    df["Label"] = [value for key, value in metadata.varLabels.items()]
    df["Measurement Level"] = [value for key, value in metadata.measureLevels.items()]
    df["Role"] = [value for key, value in metadata.varRoles.items()]
    df["Column Width"] = [value for key, value in metadata.columnWidths.items()]
    df["Alignment"] = [value for key, value in metadata.alignments.items()]
    df["Format"] = [value for key, value in metadata.formats.items()]
    frames.update({'Variable Information': df})

    # Generate Variable Values dataframe
    rows = []
    for key, value in metadata.valueLabels.items():
        i = 0
        for key2, value2 in value.items():
            if i == 0:
                rows.append([key, key2, value2])
                i = 1
            else:
                rows.append([None, key2, value2])
    df2 = pandas.DataFrame(rows)
    df2.columns = ["Variable", "Value", "Label"]
    frames.update({'Variable Values': df2})

    # Output frames to workbook
    write_frames(xlsxFile, frames)


'''
Accept command line arguments for a .sav input and a .xlsx output, or prompt the user to provide these files at the 
start or the program. This function will add file extensions to the filenames as necessary.
'''
def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--sav", type=str, help="The SPSS dataset")
    parser.add_argument("-o", "--xlsx", type=str, help="Name of XLSX to output")
    args = parser.parse_args()

    if args.sav:
        sav = args.sav
    if args.xlsx:
        xlsx = args.xlsx

    if 'sav' not in locals():
        sav = input("Name of SAV file to convert: ")
    if 'xlsx' not in locals():
        xlsx = input("Name of XLSX file to output: ")

    name, ext = os.path.splitext(sav)
    if ext == '':
        sav += '.sav'
    name, ext = os.path.splitext(xlsx)
    if ext == '':
        xlsx += '.xlsx'

    return sav, xlsx


'''
Gets header information, such as variable format, from a .sav file
* filename: the name of the .sav file
'''
def get_report(filename):
    with SavHeaderReader(filename, ioUtf8=True) as header:
        meta = header.all()
        data = str(header)
        return meta, data


'''
Writes a dictionary of pandas dataframes in {name: frame} format to an Excel Workbook (.xlsx)
* filename: the output .xlsx file
* dfs: the dictionary of pandas dataframes
'''
def write_frames(filename, dfs):
    writer = pandas.ExcelWriter(filename)
    for key, value in dfs.items():
        value.to_excel(writer, key)
    writer.save()


if __name__ == "__main__":
    main()