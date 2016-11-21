import tableauserverclient as tsc
import requests
import configparser
from Constants import *


def download(workbook, view):
    config = configparser.ConfigParser()

    config.read(SETTINGS_CONFIGURATION_FILE)
    server_name = config.get(LOGIN, SERVER)
    username = config.get(LOGIN, USERNAME)
    password = config.get(LOGIN, PASSWORD)
    site = config.get(LOGIN, SITE)

    server = tsc.Server(server_name)
    tableau_auth = tsc.TableauAuth(username, password, site)

    server.auth.sign_in(tableau_auth)

    token = server.auth_token

    url = server_name + "/t/" + site + "/views/" + workbook + "/" + view + ".pdf"
    headers = {'Cookie': 'workgroup_session_id='+token, 'Connection': 'keep-alive'}

    r = requests.get(url, headers=headers)

    if (r.content[1:4]) != b'PDF':
        raise RuntimeError("Wrong arguments, file not found. Check your workbook and sheet in settings.ini")

    with open(temp_filepath + view + '.pdf', 'wb') as fd:
        fd.write(r.content)
