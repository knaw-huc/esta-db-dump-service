from flask import Flask, request, send_file
from exporter import Exporter
app = Flask(__name__)

exporter = Exporter()





@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'POST, GET, OPTIONS'
    response.headers['Content-type'] = 'application/excel'
    return response

@app.route("/")
def hello_world():
    #exporter.make_dump()
    return send_file('static/esta.xlsx')





#Start main program

if __name__ == '__main__':
    app.run()

