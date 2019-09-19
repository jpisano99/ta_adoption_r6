from flask import Flask
from my_app import settings

app = Flask(__name__)

from my_app import views
from my_app import models
