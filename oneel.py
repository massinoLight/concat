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

    first = './Reservation_1/SMPN-RIP24-D06_05_BAKO_DISTRI.xlsx'
    df_sheet_index = pd.read_excel(first, sheet_name='Export 1')
    col = df_sheet_index.values[6]

    valeurs, commune, INSEE, path = [], [], [], []
    print(df_sheet_index.values[6][0])

    """for i in range(7, len(df_sheet_index), 4):
        valeurs.append(df_sheet_index.values[i])
        commune.append(df_sheet_index.values[1][2])
        INSEE.append(df_sheet_index.values[1][6])
        path.append(os.path.abspath(first))
    df = pd.DataFrame(valeurs, columns=col)
    df['commune']=commune
    df['INSEE'] = INSEE
    df['path'] = path
    """