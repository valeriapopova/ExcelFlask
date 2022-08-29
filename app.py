
from flask import Flask, render_template, request, redirect, Response, url_for
from werkzeug.exceptions import BadRequestKeyError

from config import Configuration
from forms import FileForm


from xls import get_data_key, get_data_value, clear_and_append, create_xls, create_worksheet

app = Flask(__name__)
app.config.from_object(Configuration)


@app.route('/', methods=['GET', 'POST'])
def homepage():
    if request.method == 'POST':
        try:
            json_file = request.form['file']
            if json_file.endswith('.json'):
                data_keys = get_data_key(json_file)
                data_values = get_data_value(json_file)
                workb = create_xls(json_file)
                worksh = create_worksheet(workb)

                clear_and_append(worksh, data_keys, data_values)

                workb.close()
                return redirect(url_for('result_page'))
            else:
                return Response("Недопустимый формат файла, необходимое расширение - JSON", 404)
        except BadRequestKeyError:
            return Response("Пустое значение", 400)

    form = FileForm()
    return render_template('homepage.html', form=form), 200


@app.route('/result')
def result_page():
    return render_template('resultpage.html'), 201
