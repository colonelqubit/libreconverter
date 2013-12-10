# LibreOffice utils.
#
# Requires Python3
#
# Based on code from:
#   PyODConverter (Python OpenDocument Converter) v1.0.0 - 2008-05-05
#   Copyright (C) 2008 Mirko Nasato <mirko@artofsolving.com>
#   Licensed under the GNU LGPL v2.1 - or any later version.
#   http://www.gnu.org/licenses/lgpl-2.1.html
#

import sys
import os
import time
import atexit


LIBREOFFICE_PORT = 8100

# Find LibreOffice.
_lopaths=(
    ('/usr/lib/libreoffice/program', '/usr/lib/libreoffice/program'),
    )

for p in _lopaths:
    if os.path.exists(p[0]):
        LIBREOFFICE_PATH    = p[0]
        LIBREOFFICE_BIN     = os.path.join(LIBREOFFICE_PATH, 'soffice')
        LIBREOFFICE_LIBPATH = p[1]

        # Add to path so we can find uno.
        if sys.path.count(LIBREOFFICE_LIBPATH) == 0:
            sys.path.insert(0, LIBREOFFICE_LIBPATH)
        break


import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException


class LORunner:
    """
    Start, stop, and connect to LibreOffice.
    """
    def __init__(self, port=LIBREOFFICE_PORT):
        """ Create LORunner that connects on the specified port. """
        self.port = port


    def connect(self, no_startup=False):
        """
        Connect to LibreOffice.
        If a connection cannot be established, try to start LibreOffice.
        """
        localContext = uno.getComponentContext()
        resolver     = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        context      = None
        did_start    = False

        n = 0
        while n < 6:
            try:
                context = resolver.resolve("uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % self.port)
                break
            except NoConnectException:
                pass

            # If first connect failed then try starting LibreOffice.
            if n == 0:
                # Exit loop if startup not desired.
                if no_startup:
                    break
                self.startup()
                did_start = True

            # Pause and try again to connect
            time.sleep(1)
            n += 1

        if not context:
            raise Exception("Failed to connect to LibreOffice on port %d" % self.port)

        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

        if not desktop:
            raise Exception("Failed to create LibreOffice desktop on port %d" % self.port)

        if did_start:
            _started_desktops[self.port] = desktop

        return desktop


    def startup(self):
        """
        Start a headless instance of LibreOffice.
        """

        args = [LIBREOFFICE_BIN,
                '--accept=socket,host=127.0.0.1,port=%d,tcpNoDelay=1;urp' % self.port,
                '--nofirststartwizard',
                '--nologo',
                '--headless',
                ]
        env  = {'PATH'       : '/bin:/usr/bin:%s' % LIBREOFFICE_PATH,
                'PYTHONPATH' : LIBREOFFICE_LIBPATH,
                }

        try:
            print("Args, Env: " + str(args) + "||" + str(env))

            pid = os.spawnve(os.P_NOWAIT, args[0], args, env)
        except Exception as e:
            raise Exception("Failed to start LibreOffice on port %d: %s" % (self.port, e.message))

        if pid <= 0:
            raise Exception("Failed to start LibreOffice on port %d" % self.port)


    def shutdown(self):
        """
        Shutdown LibreOffice.
        """
        try:
            if _started_desktops.get(self.port):
                _started_desktops[self.port].terminate()
                del _started_desktops[self.port]
        except Exception as e:
            pass



# Keep track of started desktops and shut them down on exit.
_started_desktops = {}

def _shutdown_desktops():
    """ Shutdown all LibreOffice desktops that were started by the program. """
    for port, desktop in _started_desktops.items():
        try:
            if desktop:
                desktop.terminate()
        except Exception as e:
            pass


atexit.register(_shutdown_desktops)


def lo_shutdown_if_running(port=LIBREOFFICE_PORT):
    """ Shutdown LibreOffice if it's running on the specified port. """
    lorunner = LORunner(port)
    try:
        desktop = lorunner.connect(no_startup=True)
        desktop.terminate()
    except Exception as e:
        pass


def lo_properties(**args):
    """
    Convert args to LibreOffice property values.
    """
    props = []
    for key in args:
        prop       = PropertyValue()
        prop.Name  = key
        prop.Value = args[key]
        props.append(prop)

    return tuple(props)
