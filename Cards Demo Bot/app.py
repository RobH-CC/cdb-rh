from flask import Flask, request
from webexteamssdk import WebexTeamsAPI, Webhook
from cardcontent import *
import smartsheet

app = Flask(__name__)
api = WebexTeamsAPI(access_token="NjVlMWVmNzQtOGRkZi00MTI4LThlMzgtYTEzODE2NTFkYTVmODA1OWYxZTEtZDE0_PF84_36020a6c-b90e-4d8b-a8ba-d7f9490c95dd")

@app.route('/', methods=['POST', 'GET'])
def home():
    return 'OK', 200

@app.route('/webhookreq', methods=['POST', 'GET'])
def webhookreq():
    if request.method == 'POST':
        req = request.get_json()

        data_personId = req['data']['personId']
        data_roomId = req['data']['roomId']

        #Loop prevention VERY IMPORTANT!
        me = api.people.me()
        if data_personId == me.id:
            return 'OK', 200
        else:
            if api.messages.create(roomId=data_roomId, text='Hello World!!!', attachments=[{"contentType": "application/vnd.microsoft.card.adaptive", "content": cardcontent}]):
                return "OK"

    elif request.method == 'GET':
        return "Yes, this is working."

@app.route('/cardsubmitted', methods=['POST', 'GET'])
def cardsubmitted():
    smartSheetToken = 'HXxjb5XvUfoXXMt6R87BoDGTPi0lGWbOCkyiy'
    sheetId = 2203117006677892
    nameColumnId = 1683671689258884
    emailColumnId = 6187271316629380
    phoneColumnId = 3935471502944132

    if request.method == 'POST':
        req = request.get_json()
        
        data_id = req['data']['id']
        
        attachment_actions = api.attachment_actions.get(data_id)
        inputs = attachment_actions.inputs
        myName = inputs['myName']
        myEmail = inputs['myEmail']
        myTel = inputs['myTel']
        
        print(myName)
        print(myEmail)
        print(myTel)

        smart = smartsheet.Smartsheet (smartSheetToken)
        smart.errors_as_exceptions(True)

        #Specify cell values for the added row
        newRow = smartsheet.models.Row()
        newRow.to_top = True
        newRow.cells.append({ 'column_id': nameColumnId, 'value': myName})
        newRow.cells.append({ 'column_id': emailColumnId, 'value': myEmail})
        newRow.cells.append({ 'column_id': phoneColumnId, 'value': myTel})
        response = smart.Sheets.add_rows(sheetId, newRow)

    return 'OK', 200

if __name__=='__main__':
    app.debug = True
    app.run(host='0.0.0.0')