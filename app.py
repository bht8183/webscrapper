from flask import Flask, request, send_file, render_template
from excel import create_excel
from scrappe import loop_date
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download_excel', methods=['POST'])
def download_excel():
    start_date_str = request.form['start_date']
    end_date_str = request.form['end_date']
    product = request.form['product']

    start_date = datetime.strptime(start_date_str, "%m/%d/%Y")
    end_date = datetime.strptime(end_date_str, "%m/%d/%Y")

    create_excel()
    loop_date(start_date, end_date, product)

    excel_sheet = 'output.xlsx'
    response = send_file(excel_sheet, as_attachment=True, download_name='output.xlsx')
    return response


if __name__ == '__main__':
    app.run(debug=False)