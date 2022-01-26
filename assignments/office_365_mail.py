from O365 import Account
from IPython.display import display
import ipywidgets as widgets
import codecs
import json
import random
import numpy as np
import warnings

# Deactivates deprecation warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# email template filepaths
email_template = 'resources/templates/email_template.html'
grade_email_template = 'resources/templates/grade_email_template.html'


def o365_login(id_app, secret):
    """ Function to login in Office 365 app

    Args:
        id_app (string): Azure app id
        secret (string): Azure app secret

    Returns:
        Account: Azure app handler
    """

    credentials = (id_app, secret)
    account = Account(credentials)
    account.authenticate(scopes=['basic', 'message_all'])
    return account


def generate_body(student_name, assignment):
    """ Generates the body of the email to send the assignment

    Args:
        template_path (str): location of the message template
        student_name (str): Name of the student
        assignment (Assigment): Assigment object

    Returns:
        str: email body text
    """

    f = codecs.open(email_template, 'r')
    body = f.read()

    data = assignment.config

    for i in range(len(data)):
        key = "[[" + data['Variable'][i] + "]]"
        body = body.replace(key, str(data['Value'][i]))

    return body.replace('[[name]]', student_name)


def generate_subject(assignment):
    """ Generates email subject

    Args:
        assignment (Assigment): Assigment object

    Returns:
        str: subject string
    """

    data = assignment.config

    key_1 = 'Assignment name'
    name = data.loc[assignment.config['Variable'] == key_1, 'Value']
    name = name.to_string(header=False, index=False)

    key_2 = 'Assignment code'
    code = data.loc[assignment.config['Variable'] == key_2, 'Value']

    code = code.to_string(header=False, index=False)

    return name + " - " + code


def send_email(account, assignment, email, name, attachment=False):
    """ Sends individual emails with assignment to specified student.

    Args:
        account (Account): Azure app handler
        assignment (Assigment): Assigment object
        email (str): email address
        name (str): Student name
        attachment (bool, str, optional): Attachment path.
                If False, email is sent without attachments.
                Defaults to False.

    Returns:
        bool: True if message is sent, False otherwise
    """

    m = account.new_message()
    m.to.add(email)
    m.subject = generate_subject(assignment)
    m.body = generate_body(name, assignment)
    if attachment:
        m.attachments.add(attachment)
    return m.send()


def send_email_list(account, assignment):
    """ Sends assignment emails to the student list

    Args:
        account (Account): Azure app handler
        assignment (Assigment): Assigment object

    Returns:
        bool: list with send status to all emails
    """

    data = assignment.student_list

    max_progess = len(data) - 1
    progress = widgets.FloatProgress(value=0.0, min=0.0, max=max_progess)

    print('------')
    print("Sending emails")
    display(progress)

    sent = []
    for i in range(len(data)):
        email = data['email'][i]
        name = data['name'][i]
        attachment = data['file'][i]
        sent.append(send_email(account,
                               assignment,
                               email,
                               name,
                               attachment))

        progress.value = i

    if all(sent):
        print('Emails sent with no errors')
    else:
        print('Some emails not sent (check function return')

    return sent


def load_credentials(path):
    """ Loads credentials from JSON file

    Args:
        path (str): JSON file with credentilas path

    Returns:
        dict: Dictionary with credentials
    """

    try:
        with open(path, 'r') as fp:
            return json.load(fp)
    except FileNotFoundError:
        print('Credentials file not found')
        return None


def send_test_email(account, assignment, email):
    """ Sends random email sampl to specified email address

    Args:
        account (Account): Azure app handler
        assignment (Assigment): Assigment object
        email (str): Email address
    """

    student = random.randint(0, len(assignment.student_list) - 1)
    name = assignment.student_list['name'][student]
    attachment = assignment.student_list['file'][student]

    send_email(account, assignment, email, name, attachment)


def generate_grading_table(assignment, titles, id):
    """ Generates a table with student answers, correct solutions and
        assignment points per question

    Args:
        assignment (Assigment): Assigment object
        titles (str, list): List with table titles
        id (str): Student id

    Returns:
        str: HTML formated grading table
    """

    answer_names = list(assignment.solutions.columns)[1:]
    answers = assignment.answers[assignment.answers['id'] == id][answer_names]
    if answers.empty:
        answers = ['-'] * len(answer_names)
    else:
        answers = answers.fillna('-')
        answers = answers.iloc[0].to_list()

    id_selection = assignment.student_list['id'] == id
    number = assignment.student_list[id_selection]['number'].item()

    number_selection = assignment.solutions['number'] == number
    solutions = assignment.solutions[number_selection][answer_names]

    solutions_string = []
    for solution in solutions.iloc[0].to_list():
        if np.abs(solution) >= 0.01:
            num = np.round(solution, decimals=4)
            num_format = '{}'
        else:
            num = solution
            num_format = '{:.4e}'

        solutions_string.append(num_format.format(num))

    points = assignment.grades[assignment.grades['id'] == id][answer_names]
    points = points.iloc[0].to_list()

    rows = [answers, solutions_string, points]

    table = '<center>\n <table border="1"'
    table += ' cellspacing="0" cellpadding="5" align="center">\n'
    table += '\t<tr>\n\t\t<th> </th>\n'

    for column in answer_names:
        table += '\t\t<th>{0}</th>\n'.format(column)
    table += '\t</tr>\n'

    for i, row in enumerate(rows):
        table += '\t<tr>\n\t\t<th>{0}</th>\n'.format(titles[i])
        for column in row:
            table += '\t\t<td style="text-align:center">'
            table += '{0}</td>\n'.format(column)

        table += '\t</tr>\n'

    table += '</table>\n</center>'

    return table


def send_grade_list(account, assignment, titles):
    """ Sends grade to the student list.

    Args:
        account (Account): Azure app handler.
        assignment (Assigment): Assigment object.
        titles ([str]): Grading table headers.

    Returns:
        [bool]: list with True if the email was sent, False otherwise.
    """

    grades = assignment.grades

    max_progess = len(grades) - 1
    progress = widgets.FloatProgress(value=0.0, min=0.0, max=max_progess)

    print('------')
    print("Sending emails")
    display(progress)

    sent = []

    for i, id in enumerate(assignment.grades['id']):
        sent.append(send_grade_email(account, assignment, id, titles))

        progress.value = i + 1

    if all(sent):
        print('Emails sent with no errors')
    else:
        print('Some emails not sent (check function return')

    return sent


def generate_grade_body(id, assignment, titles):
    """ Generates the body of the grade sending email

    Args:
        id (str): Student ID
        assignment (Assigment): Assigment object
        titles ([str]): Grading table headers

    Returns:
        [str]: HTML formated email body
    """

    grading_table = generate_grading_table(assignment, titles, id)

    st_list = assignment.student_list

    student_name = st_list[assignment.student_list['id'] == id]['name'].item()

    grades = assignment.grades[assignment.grades['id'] == id]

    grade = np.round(grades['grade'].values[0], decimals=1)
    points = grades['points'].values[0]

    with codecs.open(grade_email_template, 'r') as f:

        body = f.read()
        data = assignment.config

        for i in range(len(data)):
            key = "[[" + data['Variable'][i] + "]]"
            body = body.replace(key, str(data['Value'][i]))

        body = body.replace('[[name]]', student_name)
        body = body.replace('[[grade_table]]', grading_table)
        body = body.replace('[[grade]]', str(grade))
        body = body.replace('[[points]]', str(points))

    return body


def generate_grade_subject(assignment):
    """ Generates the subjet of the grade sending email

    Args:
        assignment (Assigment): Assigment object.

    Returns:
        str: Grade email subject
    """

    config = assignment.config

    name = config.loc[config['Variable'] == 'Assignment name', 'Value']
    name = name.to_string(header=False, index=False)

    code = config.loc[config['Variable'] == 'Assignment code', 'Value']
    code = code.to_string(header=False, index=False)

    return name + " - " + code


def send_grade_email(account, assignment, id, titles, email=False):
    """ Send individual grading email

    Args:
        account (Account): Azure app handler.
        assignment (Assigment): Assigment object.
        id ([type]): [description]
        titles ([str]): Grading table headers.
        email (str or bool, optional): email to send message, if set to
                                       false obtains it from assignment
                                       object. Defaults to False.

    Returns:
        [bool]: Send status.
    """

    st_list = assignment.student_list

    if not email:
        email = st_list[st_list['id'] == id]['email'].item()

    m = account.new_message()
    m.to.add(email)
    m.subject = generate_grade_subject(assignment)
    m.body = generate_grade_body(id, assignment, titles)

    return m.send()


def send_grade_test_email(account, assignment, email, titles):
    """ Sends grading email of a random ID to the specified email

    Args:
        account (Account): Azure app handler.
        assignment (Assigment): Assigment object.
        email (str): email addess
        titles ([str]): Grading table headers.

    Returns:
        [bool]: Send status.
    """
    rand_id = random.choice(assignment.grades['id'].unique())

    return send_grade_email(account, assignment, rand_id, titles, email=email)
