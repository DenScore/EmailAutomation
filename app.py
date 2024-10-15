from flask import Flask, render_template, request, render_template_string
import internalLogic

app = Flask(__name__) 



def createReturnEmail(formText):
	formDict = internalLogic.parseEmail(formText.splitlines())
	if type(formDict) == str:
		return formDict
	
	print("In createReturnEmail. Printing form dict")
	print(formDict)
	email = internalLogic.buildEmail(formDict)
	return email
	
@app.route('/', methods=['GET']) 
def index(): 
	return render_template('index.html')

@app.route('/input', methods=['POST'])
def index_post():
	returned_email=createReturnEmail(request.form['text'])
	# return render_template('results.html', returned_email=createReturnEmail(request.form['text']))
	return render_template_string(returned_email)