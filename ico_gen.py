#!/usr/bin/env python3

import PythonMagick
import sys

img = PythonMagick.Image(sys.argv[1])
img.sample('256x256')
img.write('my.ico')
