from flask import Flask, render_template,jsonify
from markit import Document

app = Flask('doccback')
# webcode = open('webcode.html').read() - not needed

@app.route('/')
def webprint():
	return render_template('home.html') 

@app.route('/blueprint', methods=['POST'])
def send_blueprint():
	
	d = Document('Mr. Naval_The Quarry Stones.pptx')

	bp = d.blueprint()

	return jsonify(bp), 200, {'ContentType':'application/json'}

if __name__ == '__main__':
	app.run(threaded=True, port = 3000,debug=True)