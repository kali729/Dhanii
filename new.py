from flask import Flask, request, jsonify, render_template, send_file
from pymongo import MongoClient
from datetime import datetime
import pandas as pd
import os
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file


from flask_pymongo import PyMongo


from werkzeug.utils import secure_filename

import urllib.parse

import os


from io import BytesIO

from flask import Flask, render_template, request, redirect, url_for, flash,session,send_file

from flask_pymongo import PyMongo

from pymongo import MongoClient

from bson.objectid import ObjectId

from werkzeug.utils import secure_filename

import os

from io import BytesIO

from flask import Flask, request, jsonify

from pymongo import MongoClient

import base64

import base64

app = Flask(__name__)

# MongoDB Configuration
username = urllib.parse.quote_plus('sftghsffth')

password = urllib.parse.quote_plus('giHkXMkhFVwBdfLb')

# Initialize MongoDB collections MongoDB Atlas connection string


app.config["MONGO_URI"] = f"mongodb+srv://{username}:{password}@cluster0.m8r2cjv.mongodb.net/dbname?retryWrites=true&w=majority"

mongo = PyMongo(app)

db = mongo.db

collection = db['clientdhani']

@app.route('/')
def home():
    return render_template("index.html")  


@app.route('/form')
def form():
    return render_template("new.html")  # Save your HTML form as `templates/form.html`

@app.route('/submit', methods=['POST'])
def submit_form():
    data = request.form.to_dict()
    data['timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Add date and time
    collection.insert_one(data)
    #return jsonify({"message": "Data saved successfully!"})
    return render_template("submited.html") 

@app.route('/download', methods=['GET'])
def download_data():
    # Fetch data from MongoDB
    records = list(collection.find({}, {"_id": 0}))  # Exclude MongoDB's ObjectId
    if not records:
        return jsonify({"message": "No data to download!"}), 404
    
    # Convert to Excel
    df = pd.DataFrame(records)
    excel_file = "loan_applications.xlsx"
    df.to_excel(excel_file, index=False)

    # Send file for download
    return send_file(excel_file, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)
