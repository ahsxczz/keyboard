#!/usr/bin/env python
# -*- coding: utf-8 -*-

from distutils.core import setup
import py2exe
import sys

if len(sys.argv) == 1:
    sys.argv.append("py2exe")
    sys.argv.append("-q")

manifest = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!-- Copyright (c) Microsoft Corporation.  All rights reserved. -->
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <noInheritable/>
    <assemblyIdentity
        type="win32"
        name="Microsoft.VC90.CRT"
        version="9.0.21022.8"
        processorArchitecture="x86"
        publicKeyToken="1fc8b3b9a1e18e3b"
    />
    <file name="msvcr90.dll" /> <file name="msvcp90.dll" /> <file name="msvcm90.dll" />
    <description>Python Interpreter</description>
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                type="win32"
                name="Microsoft.Windows.Common-Controls"
                version="6.0.0.0"
                processorArchitecture="*"
                publicKeyToken="6595b64144ccf1df"
                language="*"
            />
        </dependentAssembly>
    </dependency>
</assembly>
'''


class Target(object):
    """ A simple class that holds information on our executable file. """
    def __init__(self, **kw):
        """ Default class constructor. Update as you need. """
        self.__dict__.update(kw)

# Ok, let's explain why I am doing that.
# Often, data_files, excludes and dll_excludes (but also resources)
# can be very long list of things, and this will clutter too much
# the setup call at the end of this file. So, I put all the big lists
# here and I wrap them using the textwrap module.


data_files = []
includes = []
excludes = ['_gtkagg', 'bsddb', 'curses', 'pywin.debugger',
            'pywin.debugger.dbgcon', 'pywin.dialogs', '_hashlib',
            '_ssl', 'unittest', 'email']
# excludes = []
packages = []
dll_excludes = ['libgdk-win32-2.0-0.dll', 'libgobject-2.0-0.dll', 'tcl84.dll',
                'numpy-atlas.dll', 'tk84.dll', 'w9xpopen.exe', 'msvcp90.dll',
                "mswsock.dll", "powrprof.dll"]
# dll_excludes = []
icon_resources = [(1, 'logo.ico')]
bitmap_resources = []
other_resources = [(24, 1, manifest)]


# This is a place where the user custom code may go. You can do almost
# whatever you want, even modify the data_files, includes and friends
# here as long as they have the same variable name that the setup call
# below is expecting.

# No custom code added


# Ok, now we are going to build our target class.
# I chose this building strategy as it works perfectly for me :-D

GUI2Exe_Target_1 = Target(
    # what to build
    script="demo.py",
    icon_resources=icon_resources,
    bitmap_resources=bitmap_resources,
    other_resources=other_resources,
    dest_base=u"demo",
    version='0.1',
    company_name="Nypro & Jabil Shanghai",
    copyright="Nypro & Jabil Shanghai TE Support",
    name=u"demo",
    description='just demo',
    )

# No custom class for UPX compression or Inno Setup script

# That's serious now: we have all (or almost all) the options py2exe
# supports. I put them all even if some of them are usually defaulted
# and not used. Some of them I didn't even know about.

setup(

    # No UPX or Inno Setup

    data_files=data_files,

    options={"py2exe": {"compressed": 2,
                        "optimize": 2,
                        "includes": includes,
                        "excludes": excludes,
                        "packages": packages,
                        "dll_excludes": dll_excludes,
                        # "bundle_files": 1,
                        # 'unbuffered': 2,
                        "dist_dir": "dist",
                        "xref": False,
                        "skip_archive": False,
                        "ascii": False,
                        "custom_boot_script": '',
                        # 'typelibs': [('{C866CA3A-32F7-11D2-9602-00C04F8EE628}', 0, 5, 0)],
                        }
             },

    zipfile=None,
    console=[],
    windows=[GUI2Exe_Target_1],
    service=[],
    com_server=[],
    ctypes_com_server=[]
    )
