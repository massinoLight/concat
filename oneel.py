import os
import pandas as pd
import numpy as np


def get_file_name():
    # Use a breakpoint in the code line below to debug your script.
   return os.getcwd()


def listdirs(folder):
    return [d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))]
# Press the green button in the gutter to run the script.
if __name__ == '__main__':

        chemin='./CAP FT/Reservation_00/SMPN-RIP24-D06_04_BALA_DISTRI.xlsx'
        df_sheet_index = pd.read_excel(chemin, sheet_name='Export 1')
        #print(df_sheet_index)
        print(df_sheet_index)

        print(df_sheet_index.values[1][2])


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
