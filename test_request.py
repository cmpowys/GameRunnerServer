import requests
import zipfile
import io
import json
import cv2
import numpy as np

# payload = {
#     "commands": [ { "focus" : "Steam"}, {"wait_ms" : 2000}, {"screenshot" : "Steam"}, 
#                  {"focus" : "Chrome"}, {"wait_ms" : 2000}, {"screenshot" : "Chrome"}]
# }

# payload = {
#     "commands": [ { "focus" : "Steam"}, {"wait_ms" : 2000}, {"screenshot" : "Steam"}, {"mouse_move" : {"dx" : -11000, "dy" : 500}}]
# }

payload = {
    "commands": [ { "focus" : "notepad"}, {"wait_ms" : 2000} , {"press_keys" : "abcdefg"}]
}

response = requests.post("http://localhost:5000/command" , json=payload)

    
zip_file = zipfile.ZipFile(io.BytesIO(response.content))

response_json = zip_file.read("response.json")
results = json.loads(response_json)
filenames = [result["file"] for result in results["results"] if "file" in result]
print(results)
# for filename in filenames:
#     image_data = zip_file.read(filename)
#     nparr = np.frombuffer(image_data, np.uint8)
#     image = cv2.imdecode(nparr, flags=1)
#     cv2.imshow("TEST", image)
#     cv2.waitKey(None)