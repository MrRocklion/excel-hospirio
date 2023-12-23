
from flask import Flask,jsonify,request
from excel_generator import generarExcel
from flask_cors import CORS
app = Flask(__name__)
CORS(app)

@app.route('/excel', methods=['post'])
def query_records():
	data = dict(request.json)
	url = generarExcel(data)
	data['url'] = url
	return jsonify(data)
	
@app.route('/test', methods=['POST'])
def test():
	data = request.json
	return jsonify(data)

@app.route('/verificar_db', methods=['get'])
def verificar():
	data = request.json
	return jsonify(data)


@app.route('/')
def hello_world():
	return 'Hello World! v2'


if __name__ == "__main__":
	app.run()