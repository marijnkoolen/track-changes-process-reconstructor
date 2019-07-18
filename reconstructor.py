import xmltodict
from collections import Counter
import copy
import re

# Special keyboard presses that do not affect the text string
# or that are not used in Microsoft Word.
special_keyboard_outputs = {
    "BACK": True,
    "DELETE": True,
    "RIGHT": True,
    "LEFT": True,
    "DOWN": True,
    "UP": True,
    "LEFT Click": True,
    "TASKBAR": True,
    "LSHIFT": True,
    "END": True,
    "ESCAPE": True,
    "LCTRL + LALT": True,
    "LCTRL + LALT + @": True,
    "LALT + LCTRL": True,
    "LALT + LCTRL + @": True,
}

def interpret_event_row(event_row):
    """Cast numbers as integers and replace SPACE and RETURN by whitespace and newline characters"""
    event = copy.deepcopy(event_row)
    try:
        if "RawStart" in event:
            event["RawStart"] = int(event["RawStart"])
        if "RawEnd" in event:
            event["RawEnd"] = int(event["RawEnd"])
        event["id"] = int(event["id"])
        if "position" in event:
            event["position"] = int(event["position"])
        if "doclength" in event:
            event["doclength"] = int(event["doclength"])
        event["positionFull"] = int(event["positionFull"])
        event["doclengthFull"] = int(event["doclengthFull"])
        event["charProduction"] = int(event["charProduction"])
        if event["type"] == "keyboard" and event["output"] == "SPACE":
            event["output"] = " "
        if event["type"] == "keyboard" and event["output"] == "RETURN":
            event["output"] = "\n"
    except KeyError:
        print(event)
        raise
    return event

def make_event_list(data_json):
    """turn raw event data into parsed and interpreted event list."""
    return [interpret_event_row(event_row) for event_row in data_json['session']['event']]

def get_begin_text_string(text_begin_file):
    """read raw session text from file"""
    with open(text_begin_file, 'rt') as fh:
        begin_text_string = fh.read()
    return begin_text_string

def get_log_events(log_file):
    with open(log_file, 'rt') as fh:
        data_string = fh.read()
        data_json = xmltodict.parse(data_string)
    return make_event_list(data_json)

def update_focus(event):
    """if event is a focus event, e.g. a switch between applications or tabs,
    return which application currently has focus"""
    # focus events switch between applications
    if event["type"] == "focus" and event["output"] == "Wordlog - Microsoft Word":
        return "WORD"
    elif event["type"] == "focus" and event["output"] == "TASKBAR":
        return "TASKBAR"
    elif event["type"] == "focus" and "Windows Internet Explorer" in event["output"]:
        return "EXPLORER"
    else:
        return "UNKNOWN"

def parse_replacement(event):
    """replacement information is of the form [start_offset:end_offset] <selected_text>.
    parse replacement information into integers start and end offsets and text string of selected text"""
    match = re.match(r"\[(\d+):(\d+)\](.*)", event["output"])
    if match:
        return int(match.group(1)), int(match.group(2)), match.group(3)
    else:
        return None, None, None

def cursor_moves(event_window):
    """report if current event moves the cursor position"""
    if event_window["prev_event"]["positionFull"] != event_window["curr_event"]["positionFull"]:
        return True
    else:
        return False

def text_increases(event_window):
    """report if current event adds text"""
    if event_window["prev_event"]["doclengthFull"] < event_window["curr_event"]["doclengthFull"]:
        return True
    else:
        return False

def text_decreases(event_window):
    """report if current event removes text"""
    if event_window["prev_event"]["doclengthFull"] > event_window["curr_event"]["doclengthFull"]:
        return True
    else:
        return False

def filter_events(event_list):
    """remove all events outside of Microsoft Word, because text reconstruction only requires
    the keyboard input in Microsoft Word."""
    focus = None
    for curr_index, curr_event in enumerate(event_list):
        if curr_event["type"] == "focus":
            focus = update_focus(curr_event)
        if focus != "WORD":
            # skip everything that happens outside of Microsoft Word
            continue
        else:
            yield curr_index, curr_event

def show_event_window(event_window):
    """Rather ugly pretty printing of sliding event window."""
    print("prev_event:", event_window["prev_event"]["id"], event_window["prev_event"]["type"], event_window["prev_event"]["output"], event_window["prev_event"]["positionFull"], event_window["prev_event"]["doclengthFull"])
    print("curr_event:", curr_event["id"], curr_event["type"], curr_event["output"], curr_event["positionFull"], curr_event["doclengthFull"])
    print("next_event:", event_window["next_event"]["id"], event_window["next_event"]["type"], event_window["next_event"]["output"], event_window["next_event"]["positionFull"], event_window["next_event"]["doclengthFull"])

def slide_event_window(event_list):
    """Rather ugly way of sliding over events, keeping context of current, previous and next events."""
    event_window = {
        "prev_event": None,
        "curr_event": None,
        "next_event": None
    }
    for curr_index, curr_event in filter_events(event_list):
        event_window["curr_event"] = curr_event
        if curr_index < len(event_list) - 1:
            event_window["next_event"] = event_list[curr_index+1]
        if curr_index > 0:
            event_window["prev_event"] = event_list[curr_index-1]
        yield event_window

def is_delete(event_window):
    """report if current event is deleting (next character or selected text)"""
    if event_window["curr_event"]["type"] == "keyboard" and event_window["curr_event"]["output"] == "DELETE":
        if not is_replacement(event_window["next_event"]):
            print(event_window["curr_event"]["id"], "\tDELETE MISSES REPLACEMENT")
            return False
        else:
            return True
    return False
    return event_window["curr_event"]["type"] == "keyboard" and event_window["curr_event"]["output"] == "DELETE"

def is_backspace(event_window):
    """report if current event is backspace (removing previous character or selected text)"""
    return event_window["curr_event"]["type"] == "keyboard" and event_window["curr_event"]["output"] == "BACK"

def is_keyboard_text_remove(event_window):
    """report if current event is backspace or delete"""
    return is_delete(event_window) or is_backspace(event_window)

def is_first_output(event_window):
    """report if current event is first event that has text in Microsoft Word. This is used for skipping
    events in Word before that text is loaded."""
    if not event_window["prev_event"]:
        return False
    if event_window["prev_event"]["doclengthFull"] == 0 and event_window["curr_event"]["doclengthFull"] > 0:
        return True
    else:
        return False

def is_special_keyboard_output(event_window):
    """report if current event is a special character that doesn't affect document text."""
    if not is_keyboard_event(event_window):
        return False
    if event_window["curr_event"]["output"] in special_keyboard_outputs:
        return True
    else:
        return False

def is_keyboard_text_output(event_window):
    """report if current event is a keyboard action that normally affects document text."""
    if not is_keyboard_event(event_window):
        return False
    if is_special_keyboard_output(event_window):
        return False
    else:
        return True

def is_keyboard_event(event_window):
    """report if current event is a keyboard action"""
    if event_window["curr_event"]["type"] == "keyboard":
        return True
    else:
        return False

def is_text_load_event(event_window):
    """report if current event loads initial text in Microsoft Word"""
    if not text_increases(event_window):
        return False
    if not is_first_output(event_window):
        return False
    if is_keyboard_text_output(event_window):
        return False
    else:
        return True

def next_event_increases_text(event_window):
    """report if the next event will add to the text."""
    return event_window["next_event"]["doclengthFull"] > event_window["curr_event"]["doclengthFull"]

def next_event_decreases_text(event_window):
    """report if the next event will remove from the text."""
    return event_window["next_event"]["doclengthFull"] < event_window["curr_event"]["doclengthFull"]

def next_event_replaces_text(event_window):
    return event_window["next_event"]["type"] == "replacement"

def print_cursor_context(event_window, current_text_string, context_size):
    cursor_position = event_window["curr_event"]["positionFull"]
    curr_cursor_context = current_text_string[cursor_position-context_size:cursor_position+context_size+1]
    print(curr_cursor_context)

def get_paste_selection(event_window):
    paste_selection = event_window["next_event"]["output"]
    return re.sub(r"^\[(.*)\]$", r"\1", paste_selection)

def insert_text(event_window, current_text_string, context_size=20):
    """update the text string with the inserted keyboard input."""
    if not is_keyboard_event(event_window) and not is_paste_selection(event_window):
        print(event_window["curr_event"]["id"],"\tassume propagating correction for delayed update")
        event_window["curr_event"]["doclengthFull"] = event_window["prev_event"]["doclengthFull"]
        return current_text_string
    cursor_position = event_window["curr_event"]["positionFull"]
    insert_text = event_window["curr_event"]["output"]
    if insert_text in special_keyboard_outputs:
        insert_text = ""
    if insert_text == "LEFT + z": # HACK BASED ON SESSION 17, NEED TO CLEAN UP
        insert_text = "z"
    if is_paste_selection(event_window):
        insert_text = get_paste_selection(event_window)
        print(event_window["curr_event"]["id"], "PASTING SELECTED TEXT:", insert_text)
    before, after = current_text_string[:cursor_position], current_text_string[cursor_position:]
    next_text_string = before + insert_text + after
    if not has_expected_text_length(event_window, next_text_string):
        if event_window["curr_event"]["type"] == "keyboard" and event_window["curr_event"]["output"] == "DOWN":
            print(event_window["curr_event"]["id"], "TEXT LENGTH DISCREPANCY\texpected text length:", event_window["curr_event"]["doclengthFull"], "\tactual text length:", len(next_text_string))
            print("\tInserting newline")
            next_text_string = insert_newline(event_window, next_text_string)
        else:
            print(event_window["curr_event"]["id"], "TEXT LENGTH DISCREPANCY\texpected text length:", event_window["curr_event"]["doclengthFull"], "\tactual text length:", len(next_text_string))
            print("\t", event_window["curr_event"]["type"], "\t", event_window["curr_event"]["output"])
            next_cursor_context = next_text_string[cursor_position-context_size:cursor_position+context_size+1]
            print(next_cursor_context)
    return next_text_string

def insert_newline(event_window, curr_text_string):
    """insert a newline at the current cursor position."""
    cursor_position = event_window["curr_event"]["positionFull"]
    insert_text = "\n"
    before, after = curr_text_string[:cursor_position], curr_text_string[cursor_position:]
    return before + insert_text + after

def text_changes(event_window):
    """report if current event changes the text (even if doclengthFull is not updated properly in current event)"""
    if text_increases(event_window):
        return True
    if text_decreases(event_window):
        return True
    if is_keyboard_text_output(event_window):
        # if next event increases text length without a keyboard text input, we assume
        # curr_event is the actual text update, but the doclengthFull update is delayed
        if next_event_increases_text(event_window) and event_window["next_event"]["type"] == "keyboard" and event_window["next_event"]["output"] in special_keyboard_outputs:
            event_window["curr_event"]["doclengthFull"] = event_window["next_event"]["doclengthFull"]
            return True
        elif next_event_increases_text(event_window):
            event_window["curr_event"]["doclengthFull"] = event_window["next_event"]["doclengthFull"]
            return True
        elif event_window["next_event"]["type"] == "keyboard" and event_window["next_event"]["output"] == "DELETE" and event_window["curr_event"]["doclengthFull"] == event_window["next_event"]["doclengthFull"]:
            event_window["curr_event"]["doclengthFull"] += 1
            return True
    if is_keyboard_text_remove(event_window):
        # if curr event is text removal without decreasing text length, we assume
        # the doclengthFull update is delayed
        print(event_window["curr_event"]["id"])
        if is_backspace(event_window) and next_event_decreases_text(event_window):
            event_window["curr_event"]["doclengthFull"] -= 1
            return True
        elif is_delete(event_window) and next_event_replaces_text(event_window):
            delete_start, delete_end, delete_string = parse_replacement(event_window["next_event"])
            event_window["curr_event"]["doclengthFull"] -= delete_end - delete_start
            return True
        elif next_event_decreases_text(event_window):
            event_window["curr_event"]["doclengthFull"] = event_window["next_event"]["doclengthFull"]
            return True
    return False

def is_left_click(event):
    return event["type"] == "mouse" and event["output"] == "LEFT Click"

def is_keyboard_cut(event):
    if not event["type"] == "keyboard":
        return False
    if event["output"] == "LCTRL x":
        return True
    if event["output"] == "RCTRL x":
        return True
    else:
        return False

def is_keyboard_copy(event):
    if not event["type"] == "keyboard":
        return False
    if event["output"] == "LCTRL c":
        return True
    if event["output"] == "RCTRL c":
        return True
    else:
        return False

def is_keyboard_paste(event):
    if not event["type"] == "keyboard":
        return False
    if event["output"] == "LCTRL v":
        return True
    if event["output"] == "RCTRL v":
        return True
    else:
        return False

def is_replacement(event):
    return event["type"] == "replacement"

def is_insert(event):
    return event["type"] == "insert"

def is_cut_selection(event_window):
    if not is_left_click(event_window["curr_event"]) and not is_keyboard_cut(event_window["curr_event"]):
        return False
    if not is_replacement(event_window["prev_event"]) or not is_replacement(event_window["next_event"]):
        return False
    else:
        return True

def is_paste_selection(event_window):
    if not is_left_click(event_window["curr_event"]) and not is_keyboard_paste(event_window["curr_event"]):
        return False
    if not is_insert(event_window["next_event"]):
        return False
    else:
        print(event_window["curr_event"]["id"], "is paste selection")
        return True

def remove_text(event_window, current_text_string, context_size=20):
    """update the text string by removing selected text or previous character (backspace) or next character (delete)."""
    cursor_position = event_window["curr_event"]["positionFull"]
    if is_delete(event_window) and is_replacement(event_window["next_event"]):
        delete_start, delete_end, delete_string = parse_replacement(event_window["next_event"])
        next_text_string = current_text_string[:delete_start] + current_text_string[delete_end:]
    elif is_cut_selection(event_window):
        delete_start, delete_end, delete_string = parse_replacement(event_window["prev_event"])
        next_text_string = current_text_string[:delete_start] + current_text_string[delete_end:]
        print(event_window["curr_event"]["id"], "CUTTING SELECTED TEXT:", delete_string)
    elif is_backspace(event_window):
        delete_end = event_window["curr_event"]["positionFull"]
        delete_start = delete_end - 1
        next_text_string = current_text_string[:delete_start] + current_text_string[delete_end:]
    else:
        print(event_window)
        raise TypeError("Unknown delete sequence in event", event_window["curr_event"]["id"])
    if not has_expected_text_length(event_window, next_text_string):
        print(event_window["curr_event"]["id"], "TEXT LENGTH DISCREPANCY\texpected text length:", event_window["curr_event"]["doclengthFull"], "\tactual text length:", len(next_text_string))
        print("\t", event_window["curr_event"]["type"], "\t", event_window["curr_event"]["output"])
    return next_text_string

def update_current_text_string(event_window, current_text_string, context_size=20):
    """update the current text based on the current event. """
    if text_increases(event_window):
        return insert_text(event_window, current_text_string, context_size=context_size)
    if text_decreases(event_window):
        return remove_text(event_window, current_text_string, context_size=context_size)
    else:
        print("no change")
        raise ValueError("Update triggered with text change event for event id", event_window["curr_event"])

def has_expected_text_length(event_window, next_text_string):
    """report if text string after update of current event has the length report in the event."""
    return event_window["curr_event"]["doclengthFull"] == len(next_text_string)



