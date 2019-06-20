import urllib.request, json
import myenv
def send(dumped_json:str):
    url = myenv.URL
    method = "POST"
    headers = {"Content-Type" : "application/json"}

    # PythonオブジェクトをJSONに変換する


    # httpリクエストを準備してPOST
    request = urllib.request.Request(url, data=dumped_json, method=method, headers=headers)
    with urllib.request.urlopen(request) as response:
        response_body = response.read().decode("utf-8")
        return response_body