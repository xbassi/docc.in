from flask import Flask, render_template,jsonify,request,send_file
from markit import Document

import pprint

app = Flask('doccback')
# webcode = open('webcode.html').read() - not needed

main_doc = "Jumbo 3.pptx"

@app.route('/')
def webprint():
	return render_template('home.html') 

@app.route('/blueprint', methods=['POST'])
def send_blueprint():
	
	d = Document(main_doc)

	bp = d.blueprint()

	return jsonify(bp), 200, {'ContentType':'application/json'}

@app.route('/selectedstuff', methods=['POST'])
def get_selectedstuff():
	
	pprint.pprint(request.form)

	d = Document(main_doc)

	d.stripPPT(request.form)

	d.save("new.ppt")

	return jsonify("{'status':'success'}"),200,{'ContentType':'application/json'}


@app.route('/getfile', methods=['GET'])
def get_newfile():

	return send_file('new.ppt', attachment_filename='new.ppt',as_attachment=True)




if __name__ == '__main__':
	app.run(threaded=True, port = 3000,debug=True)