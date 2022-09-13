from flask import Flask, render_template, request, Response
from werkzeug.exceptions import BadRequestKeyError

from config import Configuration


from xls import get_data_key, get_data_value, clear_and_append, create_xls, create_worksheet

app = Flask(__name__)
app.config.from_object(Configuration)


@app.route('/excel', methods=['POST'])
def homepage():
    try:
        json_file = request.get_json(force=False)

        data_keys = get_data_key(json_file)
        data_values = get_data_value(json_file)
        workb = create_xls()
        worksh = create_worksheet(workb)

        clear_and_append(worksh, data_keys, data_values)

        workb.close()
        return Response("Проверьте таблицу", 201)

    except BadRequestKeyError:
        return Response("Пустое значение", 400)




