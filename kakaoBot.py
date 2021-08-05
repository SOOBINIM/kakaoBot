from flask import Flask, request, jsonify
import openpyxl


app = Flask(__name__)


@app.route('/', methods=['POST'])
def message():
    dataReceive = request.get_json()

    if dataReceive["userRequest"]["utterance"].startswith("예약"):
        file = openpyxl.load_workbook("예약.xlsx")
        sheet = file.active
        date = dataReceive["userRequest"]["utterance"].split(" ")[1]
        time = dataReceive["userRequest"]["utterance"].split(" ")[2]
        number = dataReceive["userRequest"]["utterance"].split(" ")[3]
        phone = dataReceive["userRequest"]["utterance"].split(" ")[4]
        i = 1
        while True:
            if sheet["A" + str(i)].value == date + " " + time:
                info = "이미 선택한 시간대에 예약이 존재합니다."
                break
            if sheet["A" + str(i)].value == None:
                sheet["A" + str(i)].value = date + " " + time
                sheet["B" + str(i)].value = number
                sheet["C" + str(i)].value = phone
                file.save("예약.xlsx")
                info = "예약되었습니다."
                break
            i += 1
        dataSend = {
            "version": "2.0",
            "template": {
                "outputs": [
                    {
                        "simpleText": {
                            "text": info
                        }
                    }
                ]
            }}

        return jsonify(dataSend)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)
