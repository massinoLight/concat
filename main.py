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
    first = './Reservation_1/SMPN-RIP24-D06_05_BAKO_DISTRI.xlsx'
    df_sheet_index = pd.read_excel(first, sheet_name='Export 1')
    col = df_sheet_index.values[6]
    valeurs, commune, INSEE, path = [], [], [], []
    for i in range(7, len(df_sheet_index), 4):
        valeurs.append(df_sheet_index.values[i])
        commune.append(df_sheet_index.values[1][2])
        INSEE.append(df_sheet_index.values[1][6])
        path.append(os.path.abspath(first))
    df = pd.DataFrame(valeurs, columns=col)
    df['commune'] = commune
    df['INSEE'] = INSEE
    df['path'] = path

    for el in os.listdir(chemin):
        valeurs, commune, INSEE, path = [], [], [], []
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


        for i in range(7, len(df_sheet_index)):
            valeurs.append(df_sheet_index.values[i])
            commune.append(df_sheet_index.values[1][2])
            INSEE.append(df_sheet_index.values[1][6])
            path.append(os.path.abspath(chemin+'/'+el+'/'+fichier))
        df2 = pd.DataFrame(valeurs, columns=col)
        df2['commune'] = commune
        df2['INSEE'] = INSEE
        df2['path'] = path
        df= df.append(df2, ignore_index=True)
    df = df[df['N° appui'].notna()]
    df.to_excel('output2.xlsx',index=False)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
