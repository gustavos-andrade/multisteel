# -*- coding: utf-8 -*-
"""
Created on Fri Mar 28 08:05:04 2025

@author: gustavo.andrade
"""

import glob
import os

# Assign directory
directory = r'Z:\ENGENHARIA\LISTA DE BOMBAS 10-2008'

# Iterate over files in directory
for filename in glob.glob(f"{directory}/*"):
    # Open file
    with open(os.path.join(directory, filename)) as f:
        print(f"Content of '{filename}'")
        # Read content of file
        print(f.read())
    print()
