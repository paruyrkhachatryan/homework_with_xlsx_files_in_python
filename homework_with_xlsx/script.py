import argparse
import xlsxwriter


def get_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--file', required=True, help="Input text file")
    parser.add_argument('-x', '--xlsxfile', required=True, help="Output Excel file")
    parser.add_argument('-s', '--sortby', choices=['Name', 'Surname', 'Age', 'Profession'], help="Sort by parameter")
    return parser.parse_args()

def get_content(filename):
    try:
        with open(filename) as f:
            return f.readlines()
    except FileNotFoundError:
        print()
    except IOError as e:
        print()

def create_person_dict(n, s, a, p):
    return {"Name": n, "Surname": s, "Age": a, "Profession": p}

def create_list_of_names(content):
    ml = []
    for line in content:
        line = line.strip()
        try:
            name, surname, age, profession = line.split()
            ml.append(create_person_dict(name, surname, age, profession))
        except ValueError:
            print()
            continue
    return ml

def sort_and_create_excel(names_list, sortby, xlsxfilename, sheetname):
    if sortby:
        try:
            sorted_list = sorted(names_list, key=lambda x: x[sortby])
        except KeyError:
            print()
    else:
        sorted_list = names_list

    workbook = xlsxwriter.Workbook(xlsxfilename)
    worksheet = workbook.add_worksheet(sheetname)
    
    cell_format = workbook.add_format({'bold': True, 'bg_color': '#90EE90'})
    ordered_list = ["Name", "Surname", "Age", "Profession"]
    
    for col, header in enumerate(ordered_list):
        worksheet.write(0, col, header, cell_format)
    
    for row, line in enumerate(sorted_list, start=1):
        for col, value in enumerate(line.values()):
            worksheet.write(row, col, value)
    
    workbook.close()



def main():
    args = get_arguments()
    fname = args.file
    xlsxfilename = args.xlsxfile
    sortby = args.sortby
    
    content = get_content(fname)
    names_list = create_list_of_names(content)
    

    sort_and_create_excel(names_list, sortby, xlsxfilename, f"Sorted by {sortby}")



if __name__ == "__main__":
    main()




