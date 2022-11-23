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
    chemin=get_file_name()+'/CAP FT'

    for el in os.listdir(chemin):
        print(el)
        files=os.listdir(chemin+'/'+el)

        x = 'SMPN'
        res = [y for y in files if x in y]

        x = '.xlsx'
        final = [y for y in res if x in y]


        if len(final)>1:
            fichier=final.__getitem__(1)
        else:
            fichier=(final.__getitem__(0))
        print(fichier)
        df_sheet_index = pd.read_excel(chemin+'/'+el+'/'+fichier, sheet_name='Export 1')
        #print(df_sheet_index)
        print(df_sheet_index.values[2][0])


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
