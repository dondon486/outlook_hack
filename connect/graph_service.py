# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
# See LICENSE in the project root for license information.

import requests
import uuid
import json
from connect.data import get_email_text
from collections import Counter
from collections import OrderedDict

# The base URL for the Microsoft Graph API.
graph_api_endpoint = 'https://graph.microsoft.com/v1.0{0}'


def call_getMails(access_token):
    # The resource URL for the sendMail action.
    send_mail_url = graph_api_endpoint.format('/me/mailFolders/SentItems/messages')

    # Set request headers.
    headers = {
        'User-Agent': 'python_tutorial/1.0',
        'Authorization': 'Bearer {0}'.format(access_token),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = {
        'client-request-id': request_id,
        'return-client-request-id': 'true'
    }
    headers.update(instrumentation)

    response = requests.get(send_mail_url, {'$select': 'toRecipients', '$top': '50'}, headers=headers, verify=False)

    list = json.loads(response.text)
    l = []
    h = {}

    for v in list.get('value'):
        recs = v.get('toRecipients')
        for r in recs:
            print
            l.append(r.get('emailAddress').get('address'))
            h[r.get('emailAddress').get('address')] = r.get('emailAddress')
    l = sort_by_freq(l)
    return [h.get(i) for i in l]

def call_getCalendarUsers(access_token):
    # The resource URL for the sendMail action.
    send_mail_url = graph_api_endpoint.format('/me/calendar/events')

    # Set request headers.
    headers = {
        'User-Agent': 'python_tutorial/1.0',
        'Authorization': 'Bearer {0}'.format(access_token),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = {
        'client-request-id': request_id,
        'return-client-request-id': 'true'
    }
    headers.update(instrumentation)

    response = requests.get(send_mail_url, {'$select': 'attendees', '$top': '50'}, headers=headers, verify=False)

    list = json.loads(response.text)
    l = []
    h = {}

    #return list
    for v in list.get('value'):
        recs = v.get('attendees')
        for r in recs:
            print
            l.append(r.get('emailAddress').get('address'))
            h[r.get('emailAddress').get('address')] = r.get('emailAddress')
    l = sort_by_freq(l)
    return [h.get(i) for i in l]

def call_getCalendarRooms(access_token):
    # The resource URL for the sendMail action.
    send_mail_url = graph_api_endpoint.format('/me/calendar/events')

    # Set request headers.
    headers = {
        'User-Agent': 'python_tutorial/1.0',
        'Authorization': 'Bearer {0}'.format(access_token),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = {
        'client-request-id': request_id,
        'return-client-request-id': 'true'
    }
    headers.update(instrumentation)

    response = requests.get(send_mail_url, {'$select': 'location', '$top': '50'}, headers=headers, verify=False)

    list = json.loads(response.text)
    l = []
    h = {}

    for v in list.get('value'):
        r = v.get('location')
        l.append(r.get('displayName'))
        h[r.get('displayName')] = r
    l = sort_by_freq(l)
    return [h.get(i) for i in l]

#top 5 sorted y frequency
def sort_by_freq(orig_list):
    final_recs = [item for items, c in Counter(orig_list).most_common() for item in [items] * c]
    final_recs = list(OrderedDict.fromkeys(final_recs))
    return final_recs[:5]


def call_sendMail_endpoint(access_token, alias, emailAddress):
    # The resource URL for the sendMail action.
    send_mail_url = graph_api_endpoint.format('/me/sendMail')

    # Set request headers.
    headers = {
        'User-Agent': 'python_tutorial/1.0',
        'Authorization': 'Bearer {0}'.format(access_token),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }

    # Use these headers to instrument calls. Makes it easier
    # to correlate requests and responses in case of problems
    # and is a recommended best practice.
    request_id = str(uuid.uuid4())
    instrumentation = {
        'client-request-id': request_id,
        'return-client-request-id': 'true'
    }
    headers.update(instrumentation)

    # Create the email that is to be sent with API.
    email = {
        'Message': {
            'Subject': 'Welcome to Office 365 development with Python and the Office 365 Connect sample',
            'Body': {
                'ContentType': 'HTML',
                'Content': get_email_text('mohit agarwal')
            },
            'ToRecipients': [
                {
                    'EmailAddress': {
                        'Address': emailAddress
                    }
                }
            ]
        },
        'SaveToSentItems': 'true'
    }

    response = requests.post(url=send_mail_url, headers=headers, data=json.dumps(email), verify=False, params=None)

    # Check if the response is 202 (success) or not (failure).
    if (response.status_code == requests.codes.accepted):
        return response.status_code
    else:
        return "{0}: {1}".format(response.status_code, response.text)
