import io
from io import BytesIO

from flask import Flask, render_template, request, Response
from openpyxl.writer.excel import save_virtual_workbook
from werkzeug.exceptions import BadRequestKeyError

from config import Configuration


from xls import get_data_key, get_data_value, clear_and_append, create_xls, create_worksheet

app = Flask(__name__)
app.config.from_object(Configuration)


@app.route('/excel/post', methods=['POST'])
def homepage():
    try:
        json_file = request.get_json(force=False)

        data_keys = get_data_key(json_file)
        data_values = get_data_value(json_file)
        workb = create_xls()
        worksh = create_worksheet(workb)

        clear_and_append(worksh, data_keys, data_values)

        workb.close()

        with open(workb.filename, "rb") as file:
            content: bytes = file.read()

        binary: str = "".join(map("{:08b}".format, content))
        print(binary)
        # content = bytes(int(binary[i: i + 8], 2) for i in range(0, len(binary), 8))
        # print(content)
        return binary

    except BadRequestKeyError:
        return Response("Пустое значение", 400)




