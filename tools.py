import re
import pandas as pd
import pathlib
from docx import Document


def round_half(number):
    round_half = number*10
    if round_half >= 35:
        return 4
    if round_half >= 25:
        return 3
    if round_half >= 15:
        return 2
    return 1


def get_match(key):
    '''
    This function get the match from the key using info.xslx
    '''
    df = pd.read_excel('./info.xlsx',
                       index_col='index',
                       sheet_name='match',)
    return df.at[key, 'info']

    def get_list_names():
        '''
    This function get the list of student names in info.xlsx
        '''
        df = pd.read_excel('./info.xlsx',
                           sheet_name='names',)
        return df.iloc[:, 0].values

def get_list_names():
    data = pd.read_excel('./info.xlsx', sheet_name='names', index_col=0)
    return data.index


def get_gender(name):
    df = pd.read_excel('./info.xlsx', index_col=0, sheet_name=2)
    return df.loc[name].iloc[0]


def get_subjects(files):
    '''
    This function create a dictionnary with the name
    of all files for a specific subject
     '''
    subject = {}
    for item in files:
        if item.is_file():
            if re.search('_', item.name):
                name = item.name.rpartition('_')
                if name[0] in subject.keys():
                    subject[name[0]].append(name[2])
                else:
                    subject[name[0]] = [name[2]]
    return subject


def get_parameter(parameter):
    ''' This function returns the value of a parameter given in arguments.
        It open infox.xlsx and the sheetnames parameters.
    '''

    parameters = pd.read_excel('./info.xlsx',
                               sheet_name='parameters',
                               index_col=0,)
    paramater = parameters.loc[parameter]
    return paramater.values[0]
