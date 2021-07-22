import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
from tkinter import *
from tkinter import filedialog

import numpy as np
import pandas as pd

import json


app = dash.Dash(suppress_callback_exceptions = True)  # Do not change that

app.layout = html.Div([html.P('Step 1: Load the dictionary'),
                       html.Button('Load dictionary from json file', id = 'load_dictionary'),
                       html.P('Step 2. If you want to update your dictionary - click the button below and provide filepath to excel file with more chinese-english translations. If not - go to step 3.'),
                       html.Button('Load excel file - to update dictionary', id = 'update_dictionary'),
                       html.P('Step 3. Load the excel file you want to translate'),
                       html.Button('Load excel file - to translate', id = 'translate_excel'),
                       dcc.Store(id = 'folder_path_dictionary'),
                       html.Div(id = 'testing'),
                        html.Div(id = 'testing2')
                      ])

@app.callback(Output('folder_path_dictionary', 'data'),
              Input('load_dictionary', 'n_clicks'))
def load_dictionary(btn1):
    value = [p['value'] for p in dash.callback_context.triggered][0]

    if value == None:
        return None

    if value > 0:

        root = Tk()

        filepath = filedialog.askopenfilename()

        try:
            root.destroy()
        except:
            pass

        return filepath


@app.callback(Output('testing', 'children'),
              Input('update_dictionary', 'n_clicks'),
              State('folder_path_dictionary', 'data'))
def update_dictionary(btn1, folder_path_dictionary):
    value = [p['value'] for p in dash.callback_context.triggered][0]

    if value == None:
        return None

    if value > 0:

        root = Tk()

        filepath_excel_to_translate = filedialog.askopenfilename()

        try:
            root.destroy()
        except:
            pass

        xl = pd.ExcelFile(filepath_excel_to_translate)

        list_of_spreadsheets = xl.sheet_names

        translation = dict()

        for i in list_of_spreadsheets:
            df = xl.parse(i)

            if '说明' in df.columns and 'Description' in df.columns:

                nans_chinese = np.where(pd.isna(df['说明']) == True)[0]
                nans_english = np.where(pd.isna(df['Description']) == True)[0]

                for jdx, j in enumerate(df['说明']):

                    if jdx not in nans_chinese:

                        if jdx not in nans_english:

                            translation[j] = df['Description'].iloc[jdx]

        try:
            dictionary = json.load(open(folder_path_dictionary, 'r'))

        except:
            dictionary = dict()

        for i in list(translation.keys()):

            if i not in list(dictionary.keys()):
                dictionary[i] = translation[i]

        json.dump(translation, open(folder_path_dictionary, 'w'))

        return None


@app.callback(Output('testing2', 'children'),
              Input('translate_excel', 'n_clicks'),
              State('folder_path_dictionary', 'data'))
def translate_excel(btn1, folder_path_dictionary):
    value = [p['value'] for p in dash.callback_context.triggered][0]

    if value == None:
        return None

    if value > 0:

        try:
            root.destroy()
        except:
            pass

        root = Tk()

        filepath_excel_to_translate = filedialog.askopenfilename()

        try:
            root.destroy()
        except:
            pass

        dictionary = json.load(open(folder_path_dictionary, 'r'))

        xl = pd.ExcelFile(filepath_excel_to_translate)

        list_of_spreadsheets = xl.sheet_names

        writer = pd.ExcelWriter(str(filepath_excel_to_translate).split('.')[0] + '- translated.xlsx', engine='xlsxwriter')

        for i in list_of_spreadsheets:

            df = xl.parse(i)

            if '说明' in df.columns and 'Description' in df.columns:

                nans_chinese = np.where(pd.isna(df['说明']) == True)[0]
                nans_english = np.where(pd.isna(df['Description']) == True)[0]

                for jdx, j in enumerate(df['说明']):

                    if jdx not in nans_chinese:

                        if jdx in nans_english:

                            if j in list(dictionary.keys()):
                                df['Description'].iloc[jdx] = dictionary[j]

            df.to_excel(writer, sheet_name=i, index=False)

        writer.save()

        writer.close()

        return None

# --- Do not change the content below
if __name__ == '__main__':
    app.run_server()