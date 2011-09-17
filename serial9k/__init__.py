#!/bin/usr/env python
# -*- coding: utf-8 -*-
import os

if os.name == 'nt':
    from serial9k.serialwin32 import *
elif os.name == 'posix':
    from serial9k.serialposix import *
else:
    raise Exception('No implementation available for \'%s\'' % os.name)
