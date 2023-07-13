
"""Module that uses sap_bot to run VBscrips, import, clean and export compiled data

"""
TEST = False
import pandas as pd
import os.path
import os
import glob
import numpy as np
import json

from vbs_scripts.CO46routine import script_CO46
from vbs_scripts.OPRRoutine import script_OPR
from vbs_scripts.EXPDRoutine import script_EXPD
from vbs_scripts.ME5Aroutine import script_ME5A
from vbs_scripts.MB25routine import script_MB25

from scripts.logger import logger
from main import ROOT_DIR


# Get Values
USER = os.environ.get('USERNAME')
projetos = pd.read_table(ROOT_DIR + "\\Project list with name.txt", header = None)
PROJECT_IDS = projetos[0].values
STD_FILES_PATH = "\\SAP data"

if TEST:
    projetos = pd.read_table(ROOT_DIR + "\\SAP data - TEST\\lista_teste.txt", header = None)
    PROJECT_IDS = projetos[0].values
    STD_FILES_PATH = "\\SAP data - TEST"

# Set values
STD_RAW_CO46 = ROOT_DIR + STD_FILES_PATH + "\\raw_co46"
STD_INT_DIR = ROOT_DIR + STD_FILES_PATH + "\\Bases"
STD_FINAL_PATH = ROOT_DIR + STD_FILES_PATH 

with open(ROOT_DIR+"\\config.json", "r") as file:
    objects = json.load(file)

COLUMNS_EXPD = objects["COLUMNS_EXPD"]
MERGED_KEYS = objects["MERGED_KEYS"]
STD_LABELS = objects["STD_LABELS"]
COLUMNS_ME5A = objects["COLUMNS_ME5A"]
COLUMNS_OPR = objects["COLUMNS_OPR"]
COLUMNS_CO46 = objects["COLUMNS_CO46"]
COLUMNS_CO46_PRE_CLEAR = objects["COLUMNS_CO46_PRE_CLEAR"]



def routine(session):
    clear_dir()
    extract_me5a(session, PROJECT_IDS, STD_INT_DIR)
    extract_expd(session, PROJECT_IDS, STD_INT_DIR)
    extract_co46_and_opr(session, PROJECT_IDS, STD_RAW_CO46, STD_INT_DIR)

    
def pandas_cleaning():
    
    #CO46
    dfs_co46 = read_files(STD_RAW_CO46, '\\CO46_', PROJECT_IDS)
    clear_co46(dfs_co46, PROJECT_IDS)
    save_df(dfs_co46, transaction = 'CO46', raw = STD_INT_DIR)
    logger.info(f"CO46 files cleaned")

    #OPR
    dfs_opr = read_files(STD_INT_DIR, '\\OPR_', PROJECT_IDS, header = [1])
    clear_opr(dfs_opr, PROJECT_IDS)
    logger.info(f"OPR files cleaned")
    
    #MERGE CO46 AND OPR
    df_co46 = concat(dfs_co46)
    df_opr = concat(dfs_opr)
    merge = merge_data(df_co46, df_opr)
    
    #CREATING FOREIGN KEYS  
    merge['FK_EXPD'] = merge['Receipt Element No.'].astype(str) + merge['Receipt Element Item'].astype(str)
    
    #REPLACE VALUES
    merge.replace(MERGED_KEYS, inplace= True)
    merge = new_dashboard_columns(merge)
    logger.info(f"Main file cleaned")
    
    #EXPD
    clear_tabs_EXPD('EXPD_complete')
    df, notnull_list = prep_dfs('EXPD_complete')
    df = concat_desc(df, notnull_list)
    df['FK_EXPD'] =  df['Purch.Doc.'].astype(str) + df['Item'].astype(str)
    save_df(df, name = 'EXPD_cleaned')
    logger.info(f"EXPD file cleaned")
    
    # ME5A
    df_me5a = cleaning_ME5A('ME5A_raw_AllProjects')
    df_me5a['FK_EXPD'] =  df_me5a['Purch.Req.'].astype(str) + df_me5a['Item'].astype(str)
    logger.info(f"ME5A file cleaned")
    
    #NEW LABELS
    replace_labels(merge, df_me5a)
    save_df(df_me5a, name = 'ME5A_cleaned')
    save_df(merge, 'cobertura_total')


def clear_dir(Temp: str = STD_RAW_CO46, bases: str = STD_INT_DIR):
    """Calls every function that is related to Directories handling.
    
    Temp: str = STD_RAW_CO46, 
        It is the standard temporary path foor every desktop
    destino: str = STD_FINAL_PATH
        Directory to clean before every run
    """
    wipe_temp(Temp)
    wipe_temp(bases)


def extract_me5a(sb, project_ids, path_intermediate:str):
    sb.run(script_ME5A, project_ids, path_intermediate)


def extract_expd(sb, project_ids, path_intermediate:str):
    sb.run(script_EXPD, project_ids, path_intermediate)


def extract_co46_and_opr(sb, project_ids, raw_co46:str, path_intermediate:str):
    for project in project_ids:
        sb.run(script_CO46, project, raw_co46)
        sb.run(script_OPR, project, path_intermediate)


def read_files( source: str,
                sufix: str,
                project_ids,
                header = None):
    """Imports all files from the suggested folder (source) with a default suffix (OPR or CO46).
    Returns a dataframes dictionary in which the keys are the project_id's
    
    source: str 
        Where to get files
    sufix: str
        What suffix to consider (OPR or CO46)
    project_ids: list like
        list os project_id's to import
    header = None  
        if the files should consider header

    The file to read is a combination of source, sufix, project_id and txt extension
    """
    df_dict = dict()
    files_encoding = 'ISO-8859-1'
    for WBS in project_ids: 
        try:
            if header == None: 
                df = pd.read_table(f'{source}{sufix}{str(WBS)}.txt', sep = '\t', encoding= files_encoding, quoting = 3)  
            else:
                df = pd.read_table(f'{source}{sufix}{str(WBS)}.txt', sep = '\t', encoding= files_encoding, header = header, quoting = 3)  
            df_dict[WBS] = df
        except:
            logger.info(f"Project {WBS} doesen't have a {sufix} file" )
    return df_dict


def clear_co46(dfs_co46, nomes_projetos):
    """Clears of all the "Unnamed" columns from dataframes
    
    """
    for project in nomes_projetos:
        try:
            colunas = list()
            df = dfs_co46[project]
            nomes_colunas = df.columns.values
            df.loc[:,'Project'] = project
            for j, name in enumerate(nomes_colunas):
                if name == str('Unnamed: ') + str(j):
                    colunas.append(name)
            df.drop(labels = colunas, axis = 1, inplace=True)
            df.drop(index = 0, axis=0, inplace = True)
            df.rename(columns={label:COLUMNS_CO46_PRE_CLEAR[i] for i,label in enumerate(df.columns)}, inplace=True)
            dfs_co46[project] = df.reindex(columns = COLUMNS_CO46)
        except:
            logger.info(f"Project {project} doesn't have an CO46 file to clear")


def concat(dfs):
    """Uses the function "concat" from pandas to concatenate dfs using key inner
    """
    concatenate = pd.concat([df for df in dfs.values()], join='inner')
    return concatenate


def clear_opr(dfs_opr, nomes_projetos):
    """Receives a dictionary containing dataframes, it will remove the column "0" of each one
    and also remove duplicates of reservation number
    """
    for name in nomes_projetos:
        try:
            df = dfs_opr[name]
            df.drop(labels='Unnamed: 0', axis=1, inplace=True)
            df.drop_duplicates(subset='Reserv.no.', inplace=True)
            df.rename(columns={label:COLUMNS_OPR[i] for i,label in enumerate(df.columns)}, inplace=True)
            df.drop(labels='Itm', axis=1, inplace=True)            
        except:
            logger.info(f"Project {name} doesen't have an OPR file to clear")


def save_df(item, name = None, transaction = None, raw:str = STD_FINAL_PATH):
    """ Saves the file (item), utilizes recursivity
        single file -> needs a name and path to directory (raw)
        dictionary  -> the file saved in raw is the concatenation of "transaction" & "key"
    """
    if type(item) == pd.DataFrame:
        item.to_csv(f'{raw}\\{name}.txt',
                    index= None, 
                    encoding='ISO-8859-1', 
                    sep= '\t', 
                    decimal=',')

    if type(item) == dict:
        for key,df in item.items():
            nome = f'{transaction}_{key}'
            save_df(df, nome, raw = raw)


def diretoriotemp(Temp):
    """Checks if there is a "Temporary" directory in the pc, if not, one will be created"""
    exists=os.path.exists(Temp)
    if not exists:
        os.mkdir(Temp)
        logger.info('Temporary path created')
    else:
        logger.info(f'Directory {Temp} already exists')
        

def wipe_temp(Temp, s_file: str = None):
    """Deletes the file(s_file) or all txt file in Temp if no file name is passed"""
    if s_file != None:
        temp_dir = f"{Temp}\\{s_file}.txt"
        os.remove(temp_dir)
        logger.info(f'File {s_file} deleted')
    else:
        temp_dir = f"{Temp}\\*.txt"
        txt_files = glob.glob(temp_dir)
        for txt_file in txt_files:
            try:
                os.remove(txt_file)
            except OSError as e:
                logger.warning(f"Error:{ e.strerror}")


def clear_tabs_EXPD(file_name:str, n_tbs:int = 9): 
    """Clear tabulation spaces that are inside commentaries of 'Expediting Description' column
        Proceeds the cleaning directly on the txt file
        
        file_name: str 
            '.txt' is not necessary, This file must be in STD_FINAL_PATH.
        n_tabs: int 
            Number o tabulation spaces to ignore in line    
    """
    with open(f'{STD_INT_DIR}\\{file_name}.txt', 'r+') as file:
        transfer_list = list()
        for line in file.readlines():
            count = 0
            line2 = ''
            for char in line:
                if char == '\t':
                    count+=1
                    if count > n_tbs:
                        char = ''
                line2 += char
            transfer_list.append(line2)
        file.truncate(0)
        file.seek(0)
        for line in transfer_list:
            file.write(line)   

#FIXME: É possível que dê erro se não houver uma PO com comentário longo, visto que não criaria automaticamente
#uma nova linha no arquivo texto e consequentemente não teria uma coluna "Unnamed :0"
#Para consertar isso é necessário checar se existem comentários quebrados antes de iniciar uma rotina de limpeza
def prep_dfs(file_name: str):
    """Imports EXPD File and return 2 objects.
        The first is the dataframe itself(EXPD), and the second is a copy of it but only of lines outside pattern 
        
        file_name_out: str
            '.txt' is not necessary, This file must be in STD_FINAL_PATH.

    """

    df = pd.read_table(f'{STD_INT_DIR}/{file_name}.txt', sep = '\t', encoding= 'ISO-8859-1', header = 1, quoting = 3)
    df.rename(columns={label:COLUMNS_EXPD[i] for i,label in enumerate(df.columns)}, inplace=True)
    long_description = df.copy(deep = True)
    long_description.dropna(axis = 0, subset = 'Unnamed: 0', inplace = True)
    long_description['Index'] = long_description.index
    long_description.reset_index(inplace = True)
    #TODO fazer list comprehension para selecionar os que quer manter e excluir colunas diferentes disso -> no caso, as colunas listadas abaixo
    #Antes de dar o drop nas colunas, concatenar o conteudo de todas elas na primeira celula, evitando que algo se perca
    columns = ['Purch.Doc.', 'Item', 'WBS element','Vendor', 'Name 1', 'Ackn. No.', 'StatDelDte', 'Deliv.date','Expediting Description']
    long_description.drop(columns = columns, inplace = True)
    notnull_list = long_description.values.tolist()
    return df, notnull_list


def concat_desc(df:pd.DataFrame, notnull_list:list):
    """Concatenates non-standard lines that have subsequent index in notnull_list. 
        Moreover, it returns the dataframe with the corrected "Expediting Description"

        df: DataFrame
            First object generated with function prep_dfs()
        notnull_list: list
            Second object generated with function prep_dfs()    
    """
    
    dict_desc = {}
    for i, line in enumerate(notnull_list):
        j = i + 1
        activate = True
        full_desc = notnull_list[i][1]
        """ Error handling for the case that 'i' is the last list position"""
        if j >= len(notnull_list): 
            activate = False
        """ Routine for identifying if there is subsequent indexes, if yes, then those lines will be concatenated"""
        while activate:
            activate = False
            try:
                next_line = notnull_list[j]
                if next_line[0] == (line[0] + j - i):
                    full_desc += ' ' + next_line[1]
                    activate = True
                    j += 1
            except:
                pass
        """Decrease one unit of index, which will be the primary key to concatenate with the original PO line
            and write in a dictonary both key and full description
        """
        desc_key = notnull_list[i][2] - 1
        dict_desc.update({desc_key:full_desc})

    """Updates all "Expediting Description" that were without a part of its original text, caused by too mny tabs """
    for i, desc in dict_desc.items():
        ind = int(i)
        concat = f"{str(df.iloc[ind]['Expediting Description'])} {str(desc)}"
        df.at[ind,'Expediting Description'] = concat
    
    df.dropna(axis=0, subset = 'Purch.Doc.', inplace=True)  
    df.drop(columns='Unnamed: 0', inplace = True)

    keys_expd = {
        "Purch.Doc." : pd.Int64Dtype(),
        "Item":pd.Int64Dtype(),
    }
    df = fix_types(df, keys_expd)
    return df


def merge_data(df1:pd.DataFrame, df2:pd.DataFrame):     
    """Receives both the CO46 and OPR compiled dataframes and merges the two using Reservation Number as key

    df1: pd.DataFrame
        CO46 Dataframe
    df2: pd.DataFrame
        OPR Dataframe
    """

    merged = df1.merge(df2, left_on='Reqmt Element Number', right_on='Reserv.no.', how='left')
    keys_merged = {
        "Receipt Element No.":pd.Int64Dtype(),
        "Receipt Element Item":pd.Int64Dtype(),
        "Reserv.no.":pd.Int64Dtype(),
        "Reqmt Element Number":pd.Int64Dtype(),
        "Pl. Deliv. Time":pd.Int64Dtype(),
        "InhseProdTime":pd.Int64Dtype()
    }
    merged = fix_types(merged, keys_merged) #converte tipos
    return merged


#cria uma coluna no dataframe com o nome do equipamento extraido da OPR
def new_dashboard_columns(merged:pd.DataFrame):
    """ Creates 2 new columns from the split of "WBS Descr." using regex: 
    "Deliverable" and "Equipment Type"
        
        merged: pandas.DataFrame
            It is the object created in the function "merge_data"

    """
    merged['WBS Descr.'].replace(to_replace='', value = '-', inplace = True, regex = True)
    df_split = merged['WBS Descr.'].str.split(pat=r"[^\w]* - ?WP?[0-9][0-9][\w]* ?- |[^\w]* - |^(?!.*- )", n=1, expand=True, regex=True)
    df_split.rename(columns={1:"Deliverable"}, inplace=True)
    df_split.drop(columns=0, inplace=True)
    df_split1 = df_split['Deliverable'].str.split(pat=r"#.*", expand=True, regex=True)
    df_split1.rename(columns={0:"Equipment Type"}, inplace=True)
    df_split1.drop(columns=1, inplace=True)
    df_split2 = merged['WBS element'].str.split(pat=r"(?<=[0-9]{6}-[\w]{4})-|(?<=[0-9]{6}-[\w]{5})-|(?<=[\w])$", n=1, expand=True, regex=True)
    df_split2.rename(columns={0:"WBS"}, inplace=True)
    df_split2.drop(columns=1, inplace=True)
    df_split3 = df_split2['WBS'].str.split(pat=r"(?<=[0-9]{6})-", n=1, expand=True, regex=True)
    df_split3.rename(columns={1:"Work Package"}, inplace=True)
    df_split3.drop(columns=0, inplace=True)
    merged = merged.join(df_split)
    merged = merged.join(df_split1)
    merged = merged.join(df_split2)
    merged = merged.join(df_split3)

    return merged 


# TODO: verificar necessidade dessa função
def fix_types(df, keys):
    return df.astype(keys)


def cleaning_ME5A(file_name:str):
    """ Imports ME5A File and return 1 object as Dataframe. 
        It drops empty columns and fixes columns data types.
        
        file_name_out: str
            '.txt' is not necessary, This file must be in STD_FINAL_PATH.

    """

    df = pd.read_table(f'{STD_INT_DIR}/{file_name}.txt', sep = '\t', encoding= 'ISO-8859-1', header = 1, quoting = 3)
    df.drop(columns='Unnamed: 0', inplace = True)
    keys_me5a = {
        "Purch.Req." : pd.Int64Dtype(),
        "Item":pd.Int64Dtype(),
        "Info rec.":pd.Int64Dtype(),
        "Reserv.no.":pd.Int64Dtype(),
        "SLoc":pd.Int64Dtype(),
        "PO":pd.Int64Dtype(),
        "Item.1":pd.Int64Dtype()
    }
    df = fix_types(df, keys_me5a)
    column_mapper = {column: COLUMNS_ME5A[i] for i, column in enumerate(df.columns)}
    df.rename(columns=column_mapper, inplace=True)
    return df


def replace_labels(df_main:pd.DataFrame, df_me5a:pd.DataFrame):
    """ Creates a new column in the DataFrame with ordered labels, 
        along with a new category for Released and NonReleased PurRqs

        df_main: must be the main dataframe, called as "merged"

        df_me5a: must be the ME5A dataframe
    """
    df_main['Receipt Element - New_label'] = df_main.loc[:, 'Receipt Element']
    df_main['Receipt Element - New_label'].replace(STD_LABELS, inplace = True)
    df_main["FK_EXPD"] = df_main["FK_EXPD"].apply(pd.to_numeric, errors = "coerce")
    df_me5a["FK_EXPD"] = df_me5a["FK_EXPD"].apply(pd.to_numeric, errors = "coerce")
    df_me5a["FK_EXPD"] = df_me5a["FK_EXPD"].astype(float)
    lista = list(df_me5a["FK_EXPD"])
    df_main.loc[df_main["Receipt Element - New_label"] == STD_LABELS['PurRqs'], "Receipt Element - New_label"] = df_main["FK_EXPD"].apply(lambda x: STD_LABELS['Handed Over PR'] if x in lista else STD_LABELS['Handover Pending'])
    df_me5a['FK_EXPD'] =  df_me5a['FK_EXPD'].astype(str)
    df_me5a['FK_EXPD'].replace(to_replace='\.0', value = '', inplace = True, regex = True)
    df_main['FK_EXPD'] =  df_main['FK_EXPD'].astype(str)
    df_main['FK_EXPD'].replace(to_replace='\.0', value = '', inplace = True, regex = True)
    

def extract_mb25(sb, path:str, plant:str = "5260"):
    sb.open_transaction("mb25")
    sb.run(script_MB25, path, "mb25", plant)
    sb.go_back()
    sb.go_back()


def clean_mb25(df_mb25):
    df_mb25["Network"] = df_mb25["Network"].astype(str)
    df_mb25["Network"] = df_mb25["Network"].str.split(pat=".", expand=True)[0]
    df_mb25["ReqmtsDate"] = pd.to_datetime(df_mb25["ReqmtsDate"], format="%d.%m.%Y").dt.strftime("%d/%m/%Y")
    df_mb25["Sales Ord."] = df_mb25["Sales Ord."].astype(str)
    df_mb25["Sales Ord."] = df_mb25["Sales Ord."].str.split(pat=".", expand=True)[0]
    df_mb25.dropna(axis="columns", how="all", inplace=True)
    return df_mb25


def get_slot_df(path):
    slot_file_names = {
                        "Mech.":"Slot dates - Setembro - Novo.xlsx",
                        "Control SCM": "Slot dates (Controls) - Setembro .xlsx",
                        "Control IMUX": "Slot dates (Controls) - Setembro .xlsx"
                    }
    df_slot = {cat: pd.read_excel(f"{path}\\{name}") for cat, name in slot_file_names.items()}
    return pd.concat(df_slot.values())
    
    
def get_db_slot_compiled(sb, plant, path:str = STD_INT_DIR):
    columns = ["Reserv.no.", "Project", "Fase", "Monday", "Friday", "Week"]
    
    extract_mb25(sb, plant)
    df_mb25 = pd.read_table(path, sep="\t", encoding="ISO-8859-1", header=1, quoting=3)
    df_mb25 = clean_mb25(df_mb25)
    
    df_slot = get_slot_df(path)

    df_merged = pd.merge(left=df_mb25, right=df_slot, on="Network", how="inner")

    df_merged["NW-At"] = df_merged.loc[:, "Network"].astype(str).str.cat(df_merged.loc[:, "OpAc"].astype(str), sep="-")
    df_merged = df_merged.loc[df_merged["Del"]!="X"]
    df_final = df_merged.drop_duplicates(subset=["NW-At"]).copy()
    df_final_ppg = df_merged.loc[df_merged["Fase"] == "PPG"].copy()
    df_final_ppg = df_final[columns]

    return df_final, df_final_ppg