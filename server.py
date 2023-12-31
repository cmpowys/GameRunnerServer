import flask
import pywinauto ## TODO use pyautogui if better?
import time
import io
import zipfile
import functools
import json
from pywinauto import mouse, keyboard
import win32api

DATE_FORMAT_STRING = "%Y/%m/%d %H:%M:%S"

app = flask.Flask(__name__)

class CommandRunner(object):
    def __init__(self):
        self.commands = dict()
    
    def run_all_commands(self, commands):
        self.file_count = 0
        response = {
            "success" : False,
            "results" : []
        }

        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as self.zip_file:
            try:
                for command_object in commands:
                    result = self.run_command(command_object)
                    response["results"].append(result)
                response["success"] = True
            except Exception as e:
                app.logger.exception(e)

            response_buffer = io.BytesIO()
            response_as_json = json.dumps(response).encode("utf-8")
            response_buffer.write(response_as_json)
            response_buffer.seek(0)
            self.zip_file.writestr("response.json", response_buffer.getvalue())

        zip_buffer.seek(0)
        return zip_buffer

    def register_command(self, function):
        command_name = function.__name__
        self.commands[command_name] = function

    def run_command(self, command_object):
        for command_name in self.commands.keys():
            if command_name in command_object:  
                command = self.commands[command_name]
                command_input = command_object[command_name]
                start_time = time.time()
                if type(command_input) is dict:
                    result = command(**command_input) or dict()
                else:
                    result = command(command_input) or dict()
                end_time = time.time()
                elapsed_ms = float(end_time - start_time) * 1000.0
                result["elapsed_ms"] = elapsed_ms

                if "file" in result:
                    filedata = result["file"]
                    filename = str(self.file_count) + ".png"
                    self.zip_file.writestr(filename, filedata.getvalue())
                    self.file_count +=1
                    result["file"] = filename
                return result
        raise Exception("Unrecognised command!")
    
command_runner = CommandRunner()

@command_runner.register_command
def press_keys(keys):
    keyboard.send_keys(keys)

@command_runner.register_command
def left_click():
    mouse.click()

@command_runner.register_command
def mouse_move(dx, dy): ## TODO broken and can not cross monitor boundaries
    cx, cy = win32api.GetCursorPos()
    mouse.move(coords=(cx + dx, cy + dy))

@command_runner.register_command
def focus(window_name):
    window = get_window(window_name)
    if window is None:
        raise Exception("Unable to focus " + window_name)
    window.set_focus()
    
@command_runner.register_command
def wait_ms(milliseconds_to_wait):
    time.sleep(float(milliseconds_to_wait) / 1000.0)

@command_runner.register_command
def screenshot(window_name):
    window = get_window(window_name)
    if window is None:
        raise Exception("Unable to take screenshot of " + window_name)   
    image = window.capture_as_image()
    image_as_bytes = io.BytesIO()
    image.save(image_as_bytes, 'PNG')
    image_as_bytes.seek(0)        
    return { "file" : image_as_bytes }

@functools.cache
def get_window(window_name):
    windows = pywinauto.Desktop().windows() ## TODO cache?
    for window in windows:
        if window_name.lower() in window.window_text().lower():
            return window  

@app.route("/command", methods=["POST"])
def command():
    commands = flask.request.get_json(force=True)["commands"]
    zip_buffer = command_runner.run_all_commands(commands)
    return flask.send_file(zip_buffer, mimetype="application/zip")

if __name__ == "__main__":
    app.run()