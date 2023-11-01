import flask
import datetime
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

commands = dict()
file_count = 0
def register_command(function):
    global commands
    command_name = function.__name__
    commands[command_name] = function

def run_command(command_object, zip_file):
    global commands
    global file_count
    for command_name in commands.keys():
        if command_name in command_object:  
            command = commands[command_name]
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
                filename = str(file_count) + ".png"
                zip_file.writestr(filename, filedata.getvalue())
                file_count +=1
                result["file"] = filename
            return result
    raise Exception("Unrecognised command!")

@register_command
def press_keys(keys):
    keyboard.send_keys(keys)

@register_command
def left_click():
    mouse.click()

@register_command
def mouse_move(dx, dy): ## TODO broken and can not cross monitor boundaries
    cx, cy = win32api.GetCursorPos()
    mouse.move(coords=(cx + dx, cy + dy))

@register_command
def focus(window_name):
    if not set_focus_to(window_name):
        raise Exception("Unable to focus " + window_name)
    
@register_command
def wait_ms(milliseconds_to_wait):
    time.sleep(float(milliseconds_to_wait) / 1000.0)

@register_command
def screenshot(window_name):
    image_as_bytes = take_screenshot(window_name)
    if (image_as_bytes is None):
        raise Exception("Unable to take screenshot of " + window_name)
    return { "file" : image_as_bytes }

@functools.cache
def get_window(window_name):
    windows = pywinauto.Desktop().windows() ## TODO cache?
    for window in windows:
        if window_name.lower() in window.window_text().lower():
            return window

def set_focus_to(window_name):
    window = get_window(window_name)
    if window is None:
        return False
    window.set_focus()
    return True

def take_screenshot(window_name):
    window = get_window(window_name)
    if window is None:
        return None
    image = window.capture_as_image()
    image_io = io.BytesIO()
    image.save(image_io, 'PNG')
    image_io.seek(0)
    return image_io

@app.route("/command", methods=["POST"])
def command():
    global file_count
    file_count = 0
    response = {
        "success" : False,
        "results" : []
    }

    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        try:
            commands = flask.request.get_json(force=True)["commands"]
            for command_object in commands:
                result = run_command(command_object, zip_file)
                response["results"].append(result)
            response["success"] = True
        except Exception as e:
            app.logger.exception(e)

        response_buffer = io.BytesIO()
        response_as_json = json.dumps(response).encode("utf-8")
        response_buffer.write(response_as_json)
        response_buffer.seek(0)
        zip_file.writestr("response.json", response_buffer.getvalue())

    zip_buffer.seek(0)
    return flask.send_file(zip_buffer, mimetype="application/zip")

if __name__ == "__main__":
    app.run()