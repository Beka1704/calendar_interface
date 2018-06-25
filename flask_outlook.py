# -*- coding: utf-8 -*-
"""
Created on Mon May 21 15:02:23 2018

@author: Benjamim
"""

from flask import Flask
from flask import request, redirect
app = Flask(__name__)

import sys, os   
sys.path.append("C:\\Users\\Benjamim\\Documents\\Python Scripts")

from urllib.parse import quote, urlencode
import base64
import json
import time
import webbrowser
import OutlookAPI

from flask import jsonify

import requests
import uuid
import json

import calendar_view_processor

# Client ID and secret
client_id = 'c8b3b2d4-f2e0-4cbd-9a02-a69648f67b30'
client_secret = 'ldndR23}%_apjUJPONQ244*'

redirect_url = 'http://127.0.0.1:5000/outlook'

# Constant strings for OAuth2 flow
# The OAuth authority
authority = 'https://login.microsoftonline.com'

# The authorize URL that initiates the OAuth2 client credential flow for admin consent
authorize_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/authorize?{0}')

# The token issuing endpoint
token_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/token')

# The scopes required by the app
scopes = [ 'openid',
           'User.Read',
           'Calendars.Read' ]


api = OutlookAPI.OutlookAPI(client_id, client_secret, redirect_url)








@app.route('/local_url', methods=['POST', 'GET'])
def outlook():
    global redirect_url, api
    error = None
    #if request.method == 'POST':
    #    print(request.form['code'])
    #    print(request.form['session_state'])
    if request.method == 'GET':
       auth_code = request.args.get('code')
       api.get_token_from_code(auth_code)
    #    print(request.form['session_state'])
    # the code below is executed if the request method
    # was GET or the credentials were invalid
    print('Token stored in API object'+api.token)
    return  'Hello Outlook code'


@app.route('/sign_benjamin_in', methods=['GET'])
def signin():
    global redirect_url
    return redirect(api.get_signin_url())

@app.route('/get_benjamin', methods=['GET'])
def getme():
    #print(api.token)
    return api.get_me()

@app.route('/get_benjamins_events', methods=['GET'])
def getmy_events():
    #print(api.token)
    if request.method == 'GET':
        date = request.args.get('date')
        events = api.get_my_events(date)
        output = calendar_view_processor.outlook_json_to_returnformat(events, date)
        resp = jsonify(output)
        resp.status_code = 200
        return resp


    #return output


app.run()   
