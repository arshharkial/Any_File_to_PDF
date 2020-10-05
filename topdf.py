import os
from os import listdir
from flask import Flask, request, redirect, url_for, render_template
from flask_mail import Mail, Message
from werkzeug.utils import secure_filename
import random
import string
import sys
import subprocess
import re
Upload_Folder = os.path.abspath(os.curdir)#path where the file is stored
app = Flask(__name__)
app.config['Upload_Folder']=Upload_Folder
app.config.update(
    DEBUG=False,
    MAIL_USERNAME ='myprojects0709@gmail.com',
    MAIL_PASSWORD ='anmoljindal@1',
    MAIL_DEFAULT_SENDER=('Anmol Jindal','myprojects0709@gmail.com'),
    MAIL_SERVER='smtp.gmail.com',
    MAIL_PORT=465,
    MAIL_USE_SSL=True
    )
mail = Mail(app)

@app.route('/', methods = ['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        email = request.form['email']
        filename=secure_filename(file.filename)
        file.save(os.path.join(app.config['Upload_Folder'],filename))
        return redirect(url_for('convert',email = email, filename = filename))
    return render_template('login.html')

def convert_to(folder, source, timeout=None):
    args = [ '--headless', '--convert-to', 'pdf', '--outdir', folder, source]

    process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    filename = re.search('-> (.*?) using filter', process.stdout.decode())

    if filename is None:
        raise LibreOfficeError(process.stdout.decode())
    else:
        return filename.group(1)

class LibreOfficeError(Exception):
    def _init_(self, output):
        self.output = output


@app.route('/uploads/<filename>/<email>')
def convert(email,filename):
    

    ifr=''.join([random.choice(string.ascii_letters+string.digits) for n in range(15)])
    os.mkdir(ifr)
    file_name = os.path.splitext(filename)[0]
    out_file = os.path.abspath(os.curdir)+'/'+ifr+'/'+file_name+'.pdf'
    
    print(filename)
    file_name=convert_to(os.path.abspath(os.curdir)+'/'+ifr,Upload_Folder+'/'+filename) 

    msg=Message('Use this id for getting your file in the future; id: '+ifr,recipients=[email])
    with app.open_resource(out_file)as fp:
        msg.attach(file_name+'.pdf', 'application/pdf',fp.read())
        mail.send(msg)
    os.remove(Upload_Folder+'/'+filename)
    return 'Your file is being processed , you will recieve the converted file on Your email id in the next 5 minutes .Your id is : '+ifr

if __name__ == '__main__':
	app.run(debug = True)
