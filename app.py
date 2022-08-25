from __future__ import print_function

from flask import Flask, render_template, redirect, flash, url_for, send_from_directory
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.csrf import CSRFProtect
from wtforms import StringField
from wtforms.validators import DataRequired, Optional, URL

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pprint import pprint
import gspread

from textblob import TextBlob

import pandas as pd

app = Flask(__name__)
app.config['SECRET_KEY'] = 'yothisissecret'
app.config['UPLOAD_FOLDER'] = 'data'
Bootstrap(app)
CSRFProtect(app)


def doc2excel(id, title, auth, pub):
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/documents.readonly']

    # The ID of a sample document.
    DOCUMENT_ID = id

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
    try:
        service = build('docs', 'v1', credentials=creds)
        # Retrieve the documents contents from the Docs service.
        document = service.documents().get(documentId=DOCUMENT_ID).execute()
        print('The title of the document is: {}'.format(document.get('title')))
        docname = document.get('title')
        print(document.get('title'))

        body = document.get('body')

        closest_htext = []
        closest_style = []
        paratext = []
        temptext = ''
        for i in range(0, len(body['content'])):
            try:
                elem = body['content'][i]['paragraph']['elements']
                for j in range(0, len(elem)):
                    try:
                        style = body['content'][i]['paragraph']['paragraphStyle']['namedStyleType']
                        content = body['content'][i]['paragraph']['elements'][j]['textRun']['content']
                        if (len(content) > 2):
                            if style == 'NORMAL_TEXT':
                                temptext = temptext + ' ' + content
                            else:
                                paratext.append(temptext)
                                temptext = ''
                                closest_htext.append(content)
                                closest_style.append(style)
                        else:
                            pass
                    except:
                        pass
            except:
                pass

        paratext = [TextBlob(text) for text in paratext]

        num_items = len(paratext)

        paratext_new = []
        closest_htext_new = []
        closest_style_new = []

        for i in range(0, num_items):
            num_sent = len(paratext[i].sentences)
            if num_sent >= 26:
                quo = num_sent // 26
                rem = num_sent % 26
                for j in range(0, quo - 1):
                    start_ind = paratext[i].sentences[j * 26].start
                    end_ind = paratext[i].sentences[(j + 1) * 26].end
                    temptext_new = paratext[i][start_ind:end_ind]
                    paratext_new.append(str(temptext_new))
                    closest_htext_new.append(closest_htext[i])
                    closest_style_new.append(closest_style[i])

                start_ind = paratext[i].sentences[(quo - 1) * 26].start
                end_ind = paratext[i].sentences[-1].end
                temptext_new = str(paratext[i][start_ind:end_ind])
                paratext_new.append(temptext_new)
                closest_htext_new.append(closest_htext[i])
                closest_style_new.append(closest_style[i])
            else:
                paratext_new.append(str(paratext[i]))
                closest_htext_new.append(closest_htext[i])
                closest_style_new.append(closest_style[i])

        data = pd.DataFrame(columns=['Title', 'Author', 'Publisher', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'Closest', 'P'])
        num_items_new = len(paratext_new)

        H1 = ''
        H2 = ''
        H3 = ''
        H4 = ''
        H5 = ''
        H6 = ''
        H7 = ''
        H8 = ''

        for i in range(0, num_items_new):

            Closest = closest_htext_new[i]
            try:
                P = paratext_new[i + 1]
            except:
                pass
            if closest_style_new[i] == 'HEADING_1':
                H1 = Closest
            elif closest_style_new[i] == 'HEADING_2':
                H2 = Closest
            elif closest_style_new[i] == 'HEADING_3':
                H3 = Closest
            elif closest_style_new[i] == 'HEADING_4':
                H4 = Closest
            elif closest_style_new[i] == 'HEADING_5':
                H5 = Closest
            elif closest_style_new[i] == 'HEADING_6':
                H6 = Closest
            elif closest_style_new[i] == 'HEADING_7':
                H7 = Closest
            elif closest_style_new[i] == 'HEADING_8':
                H8 = Closest
            else:
                pass

            entry = {'Title': title,
                     'Author': auth,
                     'Publisher': pub,
                     'H1': H1, 'H2': H2,
                     'H3': H3, 'H4': H4,
                     'H5': H5, 'H6': H6,
                     'H7': H7, 'H8': H8,
                     'Closest': Closest,
                     'P': P}

            data = data.append(entry, ignore_index=True)

            data.to_excel('static/' + docname + '.xlsx')

        return docname

    except HttpError as err:
        print(err)


class MyForm(FlaskForm):
    title = StringField('Enter Title', validators=[DataRequired()])
    author = StringField('Enter Author', validators=[DataRequired()])
    publisher = StringField('Enter Publisher', validators=[Optional()])
    link = StringField('Enter Document Link', validators=[URL()])


@app.route('/', methods=['GET', 'POST'])
def submit():
    form = MyForm()
    if form.validate_on_submit():
        act_link = form.link.data
        title = form.title.data
        author = form.author.data
        pub = form.publisher.data
        try:
            id = act_link.split('/')[5]
            print(id)
            docname = doc2excel(id, title, author, pub)
            filename = docname+'.xlsx'
        except:
            pass

        return render_template('success.html',filename=filename)
    return render_template('index.html', form=form)


if __name__ == '__main__':
    app.run(debug=True)
