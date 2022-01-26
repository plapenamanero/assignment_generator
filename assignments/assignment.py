import pandas as pd
import numpy as np
import math
import ipysheet
import ipywidgets as w
import functools
import os
import pyperclip
import warnings
from IPython.display import display
from . import gui

warnings.filterwarnings("ignore", category=DeprecationWarning)


class Assignment:

    def __init__(self, from_file=False):
        self.config = pd.DataFrame()
        self.student_list = pd.DataFrame()
        self.var_config = pd.DataFrame()
        self.variables = pd.DataFrame()
        self.solutions = pd.DataFrame()
        self.answers = pd.DataFrame()
        self.grading_config = pd.DataFrame()
        self.grades = pd.DataFrame()

        self.email_template = 'email_template.html'
        self.grades_email_template = 'grade_email_template.html'

        self._sheets = {
            'config': 'configuration',
            'student_list': 'students',
            'var_config': 'var_config',
            'variables': 'variables',
            'solutions': 'solutions',
            'answers': 'answers',
            'grading_config': "grading_config",
            'grades': 'grades'
        }

        if from_file:
            self.load_from_file()

        os.makedirs('gen', exist_ok=True)
        os.makedirs('sheets', exist_ok=True)

    def load_from_file(self):
        """ Loads attributes from XSLX file
        """

        print('------')
        try:
            fe = open('gen/data.xlsx', 'rb')
        except FileNotFoundError:
            print('Data file not found')
        else:
            print('Loading data')
            for key in self._sheets.keys():
                key_str = self._sheets[key]
                try:
                    data = pd.read_excel(fe, sheet_name=self._sheets[key])
                except ValueError as ve:
                    print(f'** {key_str} ... Data not loaded')
                    print(f'**** Error: {ve}')
                else:
                    if not data.empty:
                        sheet_name = self._sheets[key]
                        setattr(self,
                                key,
                                pd.read_excel(fe, sheet_name=sheet_name))
                        print(f'-- {key_str} ... Data loaded')
                    else:
                        print(f'-- {key_str} ... There is no data in the file')

    def configure(self):
        """ Creates Jupyter notebook interface to populate self.config

        Returns:
            ipywidget: ipywigtes layout with congiguration GUI.
        """

        columns = ['Variable', 'Value']
        variable_names = ['Greeting',
                          'Assignment name',
                          'Assignment code',
                          'Course name',
                          'Course code',
                          'Professor name',
                          'Number of questions',
                          'Number of sheets',
                          'Password']

        rows = len(variable_names)

        config_table = ipysheet.sheet(rows=rows,
                                      columns=2,
                                      column_headers=columns,
                                      row_headers=False)

        ipysheet.column(0, variable_names, read_only=True)

        if self.config.empty:
            ipysheet.column(1, [''] * rows)
        else:
            ipysheet.column(1, self.config['Value'].to_list())

        config_table.column_width = [5, 10]
        config_table.layout = w.Layout(width='500px',
                                             height='100%')
        # Creates buttons
        s_c_b_d = 'Save config'  # Save config button description
        # Save config button layout
        s_c_b_l = w.Layout(width='150px', margin='10px 0 10px 0')
        config_save_button = w.Button(description=s_c_b_d,
                                      layout=s_c_b_l)

        config_save_button.on_click(functools.partial(self.save_config,
                                                      config_table))
        # Creates interface
        gui_lay = w.Layout(margin='10px 10px 10px 10px')
        conf_gui = w.VBox([config_save_button, config_table],
                          layout=gui_lay)

        return conf_gui

    def save_config(self, config, _):
        """ Function to handle Save button in self.configure()

        Args:
            config (ipysheet table): ipysheet table with config data
            _ (): Dummy variable
        """

        self.config = ipysheet.to_dataframe(config)
        save = []

        try:
            int(self.config['Value'][6])
            save.append(True)
        except ValueError:
            print('Number of questions has to be an integer')
            save.append(False)

        try:
            int(self.config['Value'][7])
            save.append(True)
        except ValueError:
            print('Number of questions has to be an integer')
            save.append(False)

        if all(save):
            self.save_file()
            print('------')
            print('Configuration saved')
        else:
            print('------')
            print('Configuration not saved')

    def save_data(self, table, _):
        """ Saves variable configuration data

        Args:
            table (ipysheet table): ipysheet table with variable config data
            _ (): Dummy variable
        """

        self.var_config = ipysheet.to_dataframe(table)

        # Deletes lines with no variable name
        self.var_config = self.var_config[self.var_config['Variable'] != '']

        min_value = self.var_config['Min value'].astype(float)
        self.var_config['Min value'] = min_value

        max_value = self.var_config['Max value'].astype(float)
        self.var_config['Max value'] = max_value

        self.var_config['Step'] = self.var_config['Step'].astype(float)
        self.var_config['Decimals'] = self.var_config['Decimals'].astype(int)

        print('------')
        print('Variable generation configuration saved')

        self.save_file()

    def save_file(self):
        """ Saves atttribute values to XLSX file
        """

        try:
            writer = pd.ExcelWriter("gen/data.xlsx", engine='openpyxl')
        except FileNotFoundError:
            print('gen folder not found')
        else:
            print('------')
            print('Saving data file')

            for key in self._sheets.keys():
                try:
                    getattr(self, key).to_excel(writer,
                                                self._sheets[key],
                                                index=False)
                except ValueError as ve:
                    print(f'** {self._sheets[key]} ... couldn´t be saved')
                    print(f'**** Error: {ve}')
                else:
                    print(f'-- {self._sheets[key]} ... Saved')

            writer.save()
            print('------')
            print('Data saved in file')

    def load_students(self, csv=False, sep=";", auto_save=True):
        """ Loads student list from external file

        Args:
            csv (bool, optional): True if file is CSV. Defaults to False.
            sep (str, optional): Separator for CSV files. Defaults to ";".
            auto_save (bool, optional): True for save changes automatically to
                                        XLSX file. Defaults to True.
        """

        if csv:
            data_file = gui.csv_file()
            self.student_list = pd.read_csv(data_file, sep)
        else:
            data_file = gui.excel_file()
            self.student_list = pd.read_excel(data_file)

        print('------')
        print("Data loaded")

        try:
            self.add_filename()
        except KeyError:
            print('id column not found on file')

        if auto_save:
            self.save_file()
        else:
            print("Data not saved to file")

    def add_filename(self):
        """ Creates filename column in student_list DataFrame
        """

        name = self.config['Value'][2]

        sudent_list_str = self.student_list["id"].astype(str)
        files = "sheets/" + sudent_list_str + "_" + name + ".pdf"
        self.student_list['file'] = files

    def config_variables(self):
        """ Creates GUI for variable configuration (Jupyter)

        Returns:
            ipywidget: ipywidget layout
        """

        # Creates tables with legible names
        columns = ['Variable',
                   'Min value',
                   'Max value',
                   'Step',
                   'Decimals',
                   'Unit']

        table = ipysheet.sheet(rows=1,
                               columns=6,
                               column_headers=columns,
                               row_headers=False)

        if self.var_config.empty:
            table = ipysheet.sheet(rows=1,
                                   columns=6,
                                   column_headers=columns,
                                   row_headers=False)

            values = ipysheet.row(0, ['V1', 0, 0, 0, 0, ''])
        else:
            rows = len(self.var_config['Variable'])
            table = ipysheet.sheet(rows=rows,
                                   columns=6,
                                   column_headers=columns,
                                   row_headers=False)

            values = ipysheet.cell_range(self.var_config.values.tolist(),
                                         row_start=0,
                                         column_start=0)

        # Creates buttons
        add_button = w.Button(description='Add row')
        add_button.on_click(functools.partial(self.add_row, table))

        save_button = w.Button(description='Save')
        save_button.on_click(functools.partial(self.save_data, table))
        buttons_lay = w.Layout(margin='10px 0 10px ''0')
        buttons = w.HBox([add_button, save_button], layout=buttons_lay)

        # Creates gui
        gui_lay = w.Layout(margin='10px 10px 10px 10px')
        var_config_table = w.VBox([buttons, table],
                                  layout=gui_lay)

        return var_config_table

    def add_row(self, table, _):
        """ Function to handle Add button in self.config_variables()

        Args:
            table (ipysheet table): ipysheet table with config data
            _ (): Dummy variable
        """
        out = w.Output()
        with out:
            table.rows += 1
            rows_str = str(table.rows)
            ipysheet.row(table.rows - 1, ['V' + rows_str, 0, 0, 0, 0, ''])

    def generate_variable(self, low, up, step, size, decimals):
        """ Generate value for single variable

        Args:
            low (float): Minimum value of the variable.
            up (float): Maximum value of the variable.
            step (float): Difference between two consecutive generated
                          variables.
            size (int): number of steps
            decimals (int): Number of decimal positions in the generted values.

        Returns:
            float: variable value
        """

        n = math.floor(((up - low) / step) + 1)
        variable = np.random.randint(0, high=n, size=size)
        variable = variable * step
        variable = variable + low
        return np.round(variable, decimals=decimals)

    def generate_variables(self):
        """ Function to generate all random variables.
        """

        # Student data for sheet generation
        self.variables = pd.DataFrame(self.student_list['number'])
        self.variables['name'] = self.student_list['name']

        for i in range(len(self.var_config)):
            self.variables[self.var_config['Variable'][i]] = \
                self.generate_variable(self.var_config['Min value'][i],
                                       self.var_config['Max value'][i],
                                       self.var_config['Step'][i],
                                       len(self.variables),
                                       self.var_config['Decimals'][i])
        print('------')
        print('Variables generated')

        self.save_file()

    def generate_solutions(self, solver):
        """ Uses the solver() function to generate the solution list

        Args:
            solver (function): Function to solve an individual assignment.
        """

        na = self.config['Value'][6]
        self.initialize_solutions(na)

        for i in range(len(self.variables)):
            solver(self, i)

        print('------')
        print('Solutions obtained')

    def initialize_solutions(self, na):
        """ Initializes the DataFrame to store solutions

        Args:
            na (int): number for answers for each student
        """

        self.solutions = pd.DataFrame(self.variables['number'])
        # Initializes datafrane with Nan
        n = len(self.variables)
        s = np.empty([n])
        s[:] = np.nan
        for i in range(na):
            self.solutions["ap" + str(i + 1)] = s

        print('------')
        print('Solutions DataFrame initialized')

    def load_answers(self, date_format, sep=",", dec=".", auto=True):
        """ Loads students answers in a CSV format.

        Args:
            date_format (str): Date format of the CSV
            sep (str, optional): Element separator. Defaults to ",".
            dec (str, optional): Decimal separator. Defaults to ".".
            auto (bool, optional): True for automatic cleaning of answers.
                                   Defaults to True.
        """

        answers_file = gui.csv_file()
        self.answers = pd.read_csv(answers_file, sep=sep, decimal=dec)
        if auto:
            self.clean_answers_auto(date_format)

        errors = self.check_answers()

        if errors.empty:
            print('Answers loaded with no errors')
            self.save_file()
        else:
            print('Found the following errors')
            display(errors)
            print('Answers not loaded, please fix errors and try again')
            self.answers = pd.DataFrame()

    def clean_answers_auto(self, date_format):
        """ Function to clean the answers uploaded from the CSV to match the
            format required.

        Args:
            date_format (str): Date format of the CSV
        """

        ap = self.solutions.columns.tolist()
        del ap[0:1]

        columns = self.answers.columns.values.tolist()
        self.answers.drop(columns[-1], inplace=True, axis=1)
        self.answers.drop(columns[1], inplace=True, axis=1)
        columns = self.answers.columns.values.tolist()

        columns[1] = 'id'
        columns[2] = 'number'
        columns[3:] = ap[:]
        self.answers.columns = columns

        first_column = self.answers.iloc[:, 0]
        self.answers['date'] = pd.to_datetime(first_column,
                                              format=date_format)

        self.answers.sort_values(['id', 'date'], inplace=True)
        self.answers.drop_duplicates(subset=['id'], keep='last', inplace=True)
        self.answers.sort_values('number', inplace=True)

        self.answers.drop(columns[0], inplace=True, axis=1)

    def config_grading(self):
        """Creates GUI for grading configuration (Jupyter)

        Returns:
            ipywidget: ipywidget layout
        """

        ap = self.solutions.columns.tolist()
        del ap[0:1]

        ap.insert(0, 'Variable')

        grading_configuration_table = ipysheet.sheet(rows=2,
                                                     columns=len(ap),
                                                     column_headers=ap,
                                                     row_headers=False)

        ipysheet.cell(0, 0, 'Tolerance (%)', read_only=True)
        ipysheet.cell(1, 0, 'Points', read_only=True)

        if self.grading_config.empty:
            for i in range(len(ap) - 1):
                ipysheet.column(i + 1, ['', ''])
        else:
            for i in range(len(ap) - 1):
                column_values = self.grading_config.values[:, 1:][:, i]
                ipysheet.column(i + 1, column_values)

        grading_configuration_table.layout = w.Layout(width='500px',
                                                            height='100%')

        save_button = w.Button(description='Save config',
                               layout=w.Layout(width='150px',
                                               margin='10px 0 20px 0'))

        save_button.on_click(functools.partial(self.save_grading_conf,
                                               grading_configuration_table))

        gui_vBox = [save_button, grading_configuration_table]
        gui_lay = w.Layout(margin='10px 10px 10px 10px')
        grading_conf_gui = w.VBox(gui_vBox, layout=gui_lay)

        return grading_conf_gui

    def save_grading_conf(self, grading_config_table, _):
        """ Function to handle save_button in config_grading()

        Args:
            grading_config_table (ipysheet table): ipysheet table with config
                                                   data
            _ (): Dummy variable
        """

        self.grading_config = ipysheet.to_dataframe(grading_config_table)

        print('------')
        print("Configuration saved")

        self.save_file()

    def grade(self, min=0, max=10, decimals=2):
        """ Function to obtain students' grade

        Args:
            min (int, optional): Minimum grade on the scale. Defaults to 0.
            max (int, optional): Maximum grade on the scale. Defaults to 10.
            decimals (int, optional): Number of decimal on the grade.
                                      Defaults to 2.
        """

        correct = pd.DataFrame(self.student_list[['id', 'number']])

        ap = self.solutions.columns.tolist()[1:]

        tol = self.grading_config.iloc[0].to_list()[1:]
        tol = [float(item)/100 for item in tol]

        points = self.grading_config.iloc[1].to_list()[1:]
        points = [float(item) for item in points]
        tot_points = sum(points)

        for i in range(len(ap)):

            is_correct = []

            low = np.minimum(self.solutions[ap[i]] * (1 - tol[i]),
                             self.solutions[ap[i]] * (1 + tol[i]))

            up = np.maximum(self.solutions[ap[i]] * (1 - tol[i]),
                            self.solutions[ap[i]] * (1 + tol[i]))

            for student in correct['id'].to_list():

                student_number = correct[correct['id'] == student]['number']
                student_number = student_number.values[0]

                student_index = student_number - 1

                answer = self.answers[self.answers.id == student][ap[i]].values

                if answer:
                    low_value = low.iloc[student_index]
                    up_value = up.iloc[student_index]
                    is_correct.append((low_value <= answer <= up_value)[0])
                else:
                    is_correct.append(False)

            is_correct = np.array(is_correct)
            correct[ap[i]] = is_correct * points[i]

        correct.iloc[:, 2:] = correct.iloc[:, 2:].astype(float)
        correct['points'] = correct.iloc[:, 2:].sum(axis=1)

        correct['grade'] = (correct['points'] / tot_points) * (max - min) + min

        correct['grade'] = np.round(correct['grade'], decimals=1)

        self.grades = pd.DataFrame(correct)

        self.save_file()

    def get_id_string(self):
        """ Gets a regex  with the student list separated with |

        Returns:
            str: regex with the list of IDs
        """

        return '|'.join(self.student_list['id'].astype(str))

    def copy_id_string_clipboard(self):
        """ Copies regex with IDs to the clipboard
        """

        id_string = self.get_id_string()
        pyperclip.copy(id_string)
        print('String copied to clipboard')

    def set_na(self, na):
        """ Functions to change the number of answers per student

        Args:
            na (int): number of answers per student
        """

        try:
            na = int(na)
            self.config['Value'][6] = na
            self.save_file()
        except ValueError:
            print('The number of answers must be an integer')

    def set_ns(self, ns):
        """ Sets the number of pages per sheet

        Args:
            ns (int): number of pages per sheet
        """

        try:
            ns = int(ns)
            self.config['Value'][7] = ns
            self.save_file()
        except ValueError:
            print('The number of sheets must be an integer')

    def set_password(self, password):
        """ Sets password for the sheets

        Args:
            password (str): password for the sheets
        """

        self.config['Value'][8] = password
        self.save_file()

    def check_answers(self):
        """ Checks is the information provided by students in the form is
            correct

        Returns:
            bool: Pandas series with the entries wirg missmatching information
        """

        id_number_st = pd.DataFrame(self.student_list[['id', 'number']])
        cols = {'number': 'number_st'}

        id_number_st.rename(columns=cols, inplace=True)

        check_df = pd.merge(self.answers, id_number_st, on='id', how='left')
        check_df['check'] = check_df['number'] == check_df['number_st']

        return check_df[check_df['check'] == False]
