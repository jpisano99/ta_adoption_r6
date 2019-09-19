from my_app import app
from my_app.settings import app_cfg
from flask import render_template, url_for, request
import time

# from my_app.refresh_data import refresh_data, get_as_skus
# from my_app.process_bookings import process_bookings
# from my_app.build_dashboard import build_dashboard

@app.route('/')
def index():
    print('index')
    return render_template('index.html')


@app.route('/refresh_data')
def _refresh_data():
    print('refresh data')
    print(app_cfg['HOME'])
    refresh_data()
    # time.sleep(2)
    return render_template('index.html')


@app.route('/process_bookings')
def _process_bookings():
    print('process bookings')
    print(app_cfg['HOME'])
    process_bookings()
    # time.sleep(2)
    return render_template('index.html')


@app.route('/build_dashboard')
def _build_dashboard():
    print('build dashboard')
    print(app_cfg['HOME'])
    build_dashboard()
    # time.sleep(2)
    return render_template('index.html')


# @app.route('/doIt', methods=['GET', 'POST'])
# def doIt():
#     if request.method == "POST":
#         print('got a post')
#     elif request.method == "GET":
#         print('got a get')
#         print(request.args)
#
#     if request.args['action'] == 'importNew':
#         pass
#     elif request.args['action'] == 'processBookings':
#         process_bookings()
#     elif request.args['action'] == 'dashBuild':
#         build_dashboard()
#     else:
#         print('NO args found')
#
#     print(request.args['action'])
#     return render_template('status.html')


