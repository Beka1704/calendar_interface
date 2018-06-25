 # -*- coding: utf-8 -*-
"""
Created on Mon May 21 18:25:52 2018

@author: Benjamim
"""
import requests
import uuid
import json
from urllib.parse import quote, urlencode
import webbrowser
import time
from threading import Timer
from datetime import timedelta
from datetime import datetime


class OutlookAPI():
 

    client_id = None
    client_secret = None
    redirect_url = None

    expires_in = 3600

    expires_until = None
    
    # Constant strings for OAuth2 flow
    # The OAuth authority
    authority = 'https://login.microsoftonline.com'    
     # The authorize URL that initiates the OAuth2 client credential flow for admin consent
    authorize_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/authorize?{0}')
    # The token issuing endpoint
    token_url = '{0}{1}'.format(authority, '/common/oauth2/v2.0/token')
    # The scopes required by the app
    scopes = [ 'openid','User.Read','Calendars.Read','offline_access']
    
    
    token = None
    refresh_token = None
    
    graph_endpoint = 'https://graph.microsoft.com/v1.0{0}'

   
    def __init__(self, client_id, client_secret, redirect_url ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_url = redirect_url









    
    # Generic API Sending
    def make_api_call(self, method, url, token, payload = None, parameters = None):
      # Send these headers with all API calls
      headers = { 'User-Agent' : 'python_tutorial/1.0',
                  'Authorization' : 'Bearer {0}'.format(token),
                  'Accept' : 'application/json' }
    
      # Use these headers to instrument calls. Makes it easier
      # to correlate requests and responses in case of problems
      # and is a recommended best practice.
      request_id = str(uuid.uuid4())
      instrumentation = { 'client-request-id' : request_id,
                          'return-client-request-id' : 'true' }
    
      headers.update(instrumentation)
    
      response = None
      print('query url'+url)
      
      
      prepared = requests.Request(method, url, headers = headers, params = parameters).prepare()
      #self.pretty_print_Request(prepared)
    
      if (method.upper() == 'GET'):
          #print('In GET')
          response = requests.get(url, headers = headers, params = parameters)
      elif (method.upper() == 'DELETE'):
          response = requests.delete(url, headers = headers, params = parameters)
      elif (method.upper() == 'PATCH'):
          headers.update({ 'Content-Type' : 'application/json' })
          response = requests.patch(url, headers = headers, data = json.dumps(payload), params = parameters)
      elif (method.upper() == 'POST'):
          headers.update({ 'Content-Type' : 'application/json' })
          response = requests.post(url, headers = headers, data = json.dumps(payload), params = parameters)
      
     
      return response
    
    def get_signin_url(self):
      
      # Build the query parameters for the signin url
      params = { 'client_id': self.client_id,
                 'redirect_uri': self.redirect_url,
                 'response_type': 'code',
                 'scope': ' '.join(str(i) for i in self.scopes)
                }
    
      signin_url = self.authorize_url.format(urlencode(params))
    
      return signin_url
    
    def get_token_from_code(self, auth_code):
      # Build the post form for the token request
      post_data = {'grant_type': 'authorization_code',
                   'code': auth_code,
                   'redirect_uri': self.redirect_url,
                   'scope': ' '.join(str(i) for i in self.scopes),
                   'client_id': self.client_id,
                   'client_secret': self.client_secret
                   }

      now = datetime.now()
    
      r = requests.post(self.token_url, data = post_data)
      
      print(r.json())

      token = r.json()

      self.token = token['access_token']
      self.refresh_token = token['refresh_token']
      self.expires_in = token['expires_in']
      self.expires_until = now + timedelta(seconds = self.expires_in)


      # expires_in is in seconds
      # Get current timestamp (seconds since Unix Epoch) and
      # add expires_in to get expiration time
      # Subtract 5 minutes to allow for clock differences


      t = Timer(self.expires_in-300, self.get_token_from_refresh_token)
      t.start()

      print('Got and stored token!')


    def get_me(self):
        get_me_url = self.graph_endpoint.format('/me')
          # Use OData query parameters to control the results
          #  - Only return the displayName and mail fields
        query_parameters = {'$select': 'displayName,mail'}
        r = self.make_api_call('GET', get_me_url, self.token, "", parameters = query_parameters)
        print('Request Done')
        if (r.status_code == requests.codes.ok):
            print('Check done and good')
            print(type(r))
            return r.text
        else:
            print('Check done and no good')
            return "{0}: {1}".format(r.status_code, r.text)
    
    def get_my_events(self, date):

        if self.token is None:
            webbrowser.open("http://127.0.0.1:5000/sign_benjamin_in", new = 1)
            self.wait_for_token()

        if datetime.now() > self.expires_until:
            webbrowser.open("http://127.0.0.1:5000/signin", new = 1)
            self.wait_for_token()



        get_events_url = self.graph_endpoint.format('/me/calendarview')

        # Use OData query parameters to control the results
        #  - Only first 10 results returned
        #  - Only return the Subject, Start, and End fields
        #  - Sort the results by the Start field in ascending order
        # date format YYYY-MM-DDT00:00:00:00
        start = date+"T00:00:00.0"
        end = date+"T23:59:59.9"

        #"ShowAs": "Busy",
        #"IsAllDay": false,
        #"IsCancelled": false,
        query_parameters = {'$top': '100',
                            'startDateTime': start,
                            'endDateTime': end,
                            '$select': 'subject,start,end, ShowAs, IsAllDay, IsCancelled',
                            '$orderby': 'start/dateTime ASC'}

        r = self.make_api_call('GET', get_events_url, self.token, parameters = query_parameters)

        if (r.status_code == requests.codes.ok):
            return r.text
        else:
            return "{0}: {1}".format(r.status_code, "Outlook RestAPI: "+r.text)


    def wait_for_token(self):
        timeout = 180
        time_waited = 0
        while self.token is None and time_waited < timeout:
            time_waited += 3
            time.sleep(3)
        if self.token is None:
            return False
        return True

    def get_token_from_refresh_token(self):
        # Build the post form for the token request
        post_data = {'grant_type': 'refresh_token',
                     'refresh_token': self.refresh_token,
                     'redirect_uri': self.redirect_url,
                     'scope': ' '.join(str(i) for i in self.scopes),
                     'client_id': self.client_id,
                     'client_secret': self.client_secret
                     }

        r = requests.post(self.token_url, data=post_data)

        token = r.json()

        self.token = token['access_token']
        self.refresh_token = token['refresh_token']
        self.expires_in = token['expires_in']
        t = Timer(self.expires_in - 300, self.get_token_from_refresh_token)
        t.start()


