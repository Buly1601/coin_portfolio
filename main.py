from command import *
import openpyxl as pxl
import sys
import colorama
import os

def create_if_not():
    """
    Creates a new excel file if there's not an existing
    one in the folder. Name will be porfolio.xlsx
    """

    try:
        # check if file in folder
        f = open("portfolio.xlsx")

    except IOError:
        # no file in the folder, create a new one
        wb = pxl.Workbook()
        wb.save("portfolio.xlsx")

        # rename first sheet from sheet to sheet1
        wb = pxl.load_workbook("portfolio.xlsx")
        wb["Sheet"].title = "Sheet1"
        wb.save("portfolio.xlsx")

        print(colorama.Fore.MAGENTA + f"\n Created new excel file in {os.path.abspath(os.getcwd())}")
        print(colorama.Fore.RESET)


def read_from_txt(txt):
    """ 
    Receives txt file and returns list of commands
    with ammount and coin information. 
    """

    # open the file
    with open(txt) as file:
        # read each line
        lines = file.readlines()

        # check for line existance
        if not lines:
            raise Exception("No commands to execute.")

        commands = []
        # iterate through the lines and append to command list
        for line in lines:          
            content = line.split()
            
            # useful vars
            var = content[0]
            if var == "TOTAL":
                date = content[1]
            else:
                ammount = content[1]
                coin = content[2]
                if len(content) > 3:
                    mxn = content[3]
                else:
                    mxn = 0
            

            # create instance 
            instance = command(ammount, coin, mxn)

            # add vars to the commands list
            if var == "WITHDRAW":
                commands.append(instance.withdraw())
            elif var == "DEPOSIT":
                commands.append(instance.deposit())
            elif var == "TOTAL":
                commands.append(instance.total(date))
            # elif var == "GRAPH":
            #     commands.append(instance.graph())
            else:
                raise Exception(f"Invalid command '{var}'")

        return commands


def main():
    
    # check if command-line
    if len(sys.argv) != 2:
        raise Exception(f"USAGE: {__file__} .txt_file") # TODO filename

    # read text file
    txt_file = sys.argv[1]

    # check for needed elements
    create_if_not()

    # get commands
    commands = read_from_txt(txt_file)
    
    # read from command list
    for c in commands:
        # call command
        c


if __name__ == "__main__":
    main()