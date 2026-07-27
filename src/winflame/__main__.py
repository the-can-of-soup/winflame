"""
Command-line utility for generating flame graphs from file trees on Windows.
"""



# IMPORTS

from . import cli
import sys



# MAIN

if __name__ == '__main__':
    sys.exit(cli.cli())
