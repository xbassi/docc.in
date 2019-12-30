from flask import Flask, render_template

app = Flask('doccback')
# webcode = open('webcode.html').read() - not needed

@app.route('/')
def webprint():
    return render_template('home.html') 

if __name__ == '__main__':
    app.run(threaded=True, port = 3000)