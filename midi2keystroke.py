#!/usr/bin/env python
#
# midi2keystroke.py
#
"""Execute external commands when specific MIDI messages are received.

Example configuration (in YAML syntax)::

    - name: My Backingtracks
      description: Play audio file with filename matching <data1>-playback.mp3
        when program change on channel 16 is received
      status: programchange
      channel: 16
      command: plaympeg %(data1)03i-playback.mp3
    - name: My Lead Sheets
      description: Open PDF with filename matching <data2>-sheet.pdf
        when control change 14 on channel 16 is received
      status: controllerchange
      channel: 16
      data: 14
      command: evince %(data2)03i-sheet.pdf

"""

import argparse
import logging
import sys
import time

from os.path import exists

import sendKey

try:
    from functools import lru_cache
except ImportError:
    # Python < 3.2
    try:
        from backports.functools_lru_cache import lru_cache
    except ImportError:
        lru_cache = lambda: lambda func: func

import yaml

import rtmidi
from rtmidi.midiutil import open_midiinput
from rtmidi.midiconstants import (CHANNEL_PRESSURE, CONTROLLER_CHANGE, NOTE_ON, NOTE_OFF,
                                  PITCH_BEND, POLY_PRESSURE, PROGRAM_CHANGE)

log = logging.getLogger('midi2command')
BACKEND_MAP = {
    'alsa': rtmidi.API_LINUX_ALSA,
    'jack': rtmidi.API_UNIX_JACK,
    'coremidi': rtmidi.API_MACOSX_CORE,
    'windowsmm': rtmidi.API_WINDOWS_MM
}
STATUS_MAP = {
    'noteon': NOTE_ON,
    'noteoff': NOTE_OFF,
    'programchange': PROGRAM_CHANGE,
    'controllerchange': CONTROLLER_CHANGE,
    'pitchbend': PITCH_BEND,
    'polypressure': POLY_PRESSURE,
    'channelpressure': CHANNEL_PRESSURE
}

KEY_UP = 1
KEY_DOWN = 2
KEY_DOWN_UP = KEY_UP | KEY_DOWN


class KeyStroke(object):
    def __init__(self, name='', description='', status=0xB0, channel=None, data=None,
                 keys=None):
        self.name = name
        self.description = description
        self.status = status
        self.channel = channel
        self.keys = keys.split()

        if data is None or isinstance(data, int):
            self.data = data
        elif hasattr(data, 'split'):
            self.data = [int(n) for n in data.split()]
        else:
            raise TypeError("Could not parse 'data' field.")


class MidiInputHandler(object):
    def __init__(self, port, config):
        self.port = port
        self._wallclock = time.time()
        self.keystrokes = dict()
        self.load_config(config)

    def __call__(self, event, data=None):
        event, deltatime = event
        self._wallclock += deltatime

        if event[0] < 0xF0:
            channel = (event[0] & 0xF) + 1
            status = event[0] & 0xF0
        else:
            status = event[0]
            channel = None

        data1 = data2 = None
        num_bytes = len(event)

        if num_bytes >= 2:
            data1 = event[1]
        if num_bytes >= 3:
            data2 = event[2]

        log.debug("[%s] @%i CH:%2s s:%02X d1:%s d2:%s", self.port, self._wallclock,
                  channel or '-', status, data1, data2 or '')

        # Look for matching command definitions
        cmd = self.lookup_command(status, channel, data1, data2)

        if cmd:
            action_type = KEY_DOWN
            if status == NOTE_OFF:
                action_type = KEY_UP
            if status == CONTROLLER_CHANGE:
                action_type = KEY_DOWN_UP

            self.do_command(cmd.keys, action_type)
        else:
            log.debug("no cmd")

    @lru_cache
    def lookup_command(self, status, channel, data1, data2):
        if status == NOTE_OFF:
            status = NOTE_ON
        elif status == CONTROLLER_CHANGE and data2 < 64:
            data2 = 63
        elif status == CONTROLLER_CHANGE and data2 > 64:
            data2 = 65

        for keystroke in self.keystrokes.get(status, []):
            if channel is not None and keystroke.channel != channel:
                continue

            if status == NOTE_ON and keystroke.data == data1:
                return keystroke
            elif (isinstance(keystroke.data, list) and
                  keystroke.data[0] == data1 and keystroke.data[1] == data2):
                return keystroke

    @staticmethod
    def do_command(keystrokes, action_type):
        try:
            if action_type == KEY_DOWN or action_type == KEY_DOWN_UP:
                for keystroke in keystrokes:
                    log.info("press: %s", keystroke)
                    keycode = sendKey.SetKeyboardConsts(keystroke)
                    sendKey.PressKey(keycode)

            if action_type == KEY_DOWN_UP:
                log.info("delay")
                time.sleep(0.01)

            if action_type == KEY_UP or action_type == KEY_DOWN_UP:
                for keystroke in keystrokes:
                    log.info("release: %s", keystroke)
                    keycode = sendKey.SetKeyboardConsts(keystroke)
                    sendKey.ReleaseKey(keycode)
        except:  # noqa: E722
            log.exception("Error calling external command.")

    def load_config(self, filename):
        if not exists(filename):
            raise IOError("Config file not found: %s" % filename)

        with open(filename) as patch:
            data = yaml.load(patch, Loader=yaml.FullLoader)

        for cmdspec in data:
            try:
                if isinstance(cmdspec, dict) and 'keys' in cmdspec:
                    cmd = KeyStroke(**cmdspec)
                elif len(cmdspec) >= 2:
                    cmd = KeyStroke(*cmdspec)
            except (TypeError, ValueError) as exc:
                log.debug(cmdspec)
                raise IOError("Invalid command specification: %s" % exc)
            else:
                status = STATUS_MAP.get(cmd.status.strip().lower())

                if status is None:
                    try:
                        int(cmd.status)
                    except:  # noqa: E722
                        log.error("Unknown status '%s'. Ignoring command",
                                  cmd.status)

                log.debug("Config: %s\n%s\n%s\n", cmd.name, cmd.description, cmd.keys)
                self.keystrokes.setdefault(status, []).append(cmd)


def main(args=None):
    """Main program function.

    Parses command line (parsed via ``args`` or from ``sys.argv``), detects
    and optionally lists MIDI input ports, opens given MIDI input port,
    and attaches MIDI input handler object.

    """
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    padd = parser.add_argument
    padd('-b', '--backend', choices=sorted(BACKEND_MAP),
         help='MIDI backend API (default: OS dependent)')
    padd('-p', '--port',
         help='MIDI input port name or number (default: open virtual input)')
    padd('-v', '--verbose',
         action="store_true", help='verbose output')
    padd(dest='config', metavar="CONFIG",
         help='Configuration file in YAML syntax.')

    args = parser.parse_args(args)

    logging.basicConfig(format="%(name)s: %(levelname)s - %(message)s",
                        level=logging.DEBUG if args.verbose else logging.INFO)

    try:
        midiin, port_name = open_midiinput(
            args.port,
            use_virtual=True,
            api=BACKEND_MAP.get(args.backend, rtmidi.API_UNSPECIFIED),
            client_name='midi2command',
            port_name='MIDI input')
    except (IOError, ValueError) as exc:
        return "Could not open MIDI input: %s" % exc
    except (EOFError, KeyboardInterrupt):
        return

    log.debug("Attaching MIDI input callback handler.")
    midiin.set_callback(MidiInputHandler(port_name, args.config))

    log.info("Entering main loop. Press Control-C to exit.")
    try:
        # just wait for keyboard interrupt in main thread
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print('')
    finally:
        midiin.close_port()
        del midiin


if __name__ == '__main__':
    sys.exit(main() or 0)
