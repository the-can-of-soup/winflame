# IMPORTS

from __future__ import annotations

from PIL import Image, ImageFont, PcfFontFile, BdfFontFile
from collections.abc import Callable
from typing import Any, NoReturn
import win32com.shell.shell
import pickletools
import platform
import argparse
import hashlib
import pickle
import shutil
import time
import sys
import os

import winflame
import common



# CONSTANTS

# Relative to program directory
CACHE_DIR: str = 'cache'
OUTPUT_DIR: str = 'output'
TREE_FILE_HASHES_FILENAME: str = '.tree_file_hashes'

# Relative to cache directory
CACHED_FILE_TREE_FILENAME: str = 'file_tree.wftree'

# Relative to output directory, in `strftime` format
DEFAULT_FLAME_GRAPH_FILENAME_FORMAT: str = 'flame_graph_%H-%M-%S_%m-%d-%Y.png'
DEFAULT_FILE_TREE_FILENAME_FORMAT: str = 'file_tree_%H-%M-%S_%m-%d-%Y.wftree'

# CLI colors
SUCCESS_COLOR: tuple[int, int, int] = (50, 255, 50)
WARNING_COLOR: tuple[int, int, int] = (255, 255, 0)
ERROR_COLOR: tuple[int, int, int] = (255, 50, 50)

COLOR_KEY: list[tuple[str, tuple[int, int, int]] | None] = [ # Each list element is a line; `None` indicates a blank line
    ('Regular file*', winflame.FileType.REGULAR_FILE.color_rgb),
    ('Directory*', winflame.FileType.DIRECTORY.color_rgb),
    ('Alternate data stream*', winflame.FileType.ALTERNATE_DATA_STREAM.color_rgb),
    ('Symlink', winflame.FileType.SYMLINK.color_rgb),
    ('Junction', winflame.FileType.JUNCTION.color_rgb),
    ('Hard link (new name for a previously-seen file)', winflame.FileType.HARD_LINK.color_rgb),
    ('Unknown file type', winflame.FileType.UNKNOWN.color_rgb),
    None,
    ('Processing not started', winflame.NOT_STARTED_COLOR),
    ('Error during processing', winflame.ERROR_COLOR),
    None,
    ('Drive unaccounted space*', winflame.UNACCOUNTED_COLOR[:3]),
    ('Drive free space*', winflame.FREE_COLOR[:3]),
    ('Drive extra-counted space*', winflame.EXTRA_COUNTED_COLOR[:3]),
]



# GLOBALS

silent_mode: bool = False # Suppresses all output except errors and warnings
suppress_warnings: bool = False



# DEFINITIONS

def load_file_tree(file_path: str, tree_file_hashes_path: str) -> winflame.FileNode | None:
    """
    Loads a file tree from a file and verifies that the file was created by this program installation.

    :param file_path: The path to load the file tree from.
    :type file_path: str
    :param tree_file_hashes_path: The path of the tree file hashes file (does not need to exist yet).
    :type tree_file_hashes_path: str
    :return: The root node of the file tree, or ``None`` if the security check failed.
    :rtype: winflame.FileNode | None
    """
    # The file is guaranteed to fail the security test because we do not have a trusted hashes file yet
    if not os.path.exists(tree_file_hashes_path):
        return None

    # Compute file hash digest
    with open(file_path, 'rb') as f:
        digest: bytes = hashlib.file_digest(f, hashlib.sha256).digest()

    # Compare against previously-created file hashes
    with open(tree_file_hashes_path, 'rb') as f:
        while True:
            # Read a trusted digest (32 bytes for SHA256)
            trusted_digest: bytes = f.read(32)
            if len(trusted_digest) < 32:
                # EOF
                return None

            # Check for match
            if digest == trusted_digest:
                break

    # If we reach this point, the file has been verified as safe.

    # Load and unpickle file tree
    with open(file_path, 'rb') as f:
        return pickle.load(f)

def save_file_tree(root: winflame.FileNode, file_path: str, tree_file_hashes_path: str) -> None:
    """
    Saves a file tree to a file and saves its hash digest so it can be loaded later with ``load_file_tree``.

    :param root: The root node of the file tree.
    :type root: winflame.FileNode
    :param file_path: The path to save the file tree to.
    :type file_path: str
    :param tree_file_hashes_path: The path of the tree file hashes file (does not need to exist yet).
    :type tree_file_hashes_path: str
    """
    # Pickle file tree
    pickled: bytes = pickletools.optimize(pickle.dumps(root))

    # Compute digest
    digest: bytes = hashlib.sha256(pickled).digest()

    # Save file tree to file
    with open(file_path, 'wb') as f:
        f.write(pickled)

    # Save digest to trusted digest list
    with open(tree_file_hashes_path, 'a+b') as f:
        # Check if digest is already in trusted digest list
        f.seek(0)
        while True:
            # Read a trusted digest (32 bytes for SHA256)
            trusted_digest: bytes = f.read(32)
            if len(trusted_digest) < 32:
                # EOF
                break

            # Check for match
            if digest == trusted_digest:
                # Already in trusted digest list; our work here is done
                return

        # If we reach this point, the hash digest is not already in the trusted list.

        # Add digest to trusted digest list
        f.write(digest)

def generic_message(
        message: str,
        end: str = '\n',
        ask_yes_no: bool = False,
        clear_line: bool = True,
        bypass_silent_mode: bool = False,
    ) -> bool | None:
    """
    Prints a message.

    :param message: The message to print.
    :type message: str
    :param end: A string to append to the message before printing.
    :type end: str
    :param ask_yes_no: If ``True``, the user will be prompted for a yes/no answer (defaulting to "no" if invalid input,
        or "yes" if suppressed), and ``end`` will have no effect. Terminal formatting will automatically be cleared at
        the end of the line if this is ``True``.
    :type ask_yes_no: bool
    :param clear_line: Whether to clear all characters after the cursor on the current line before printing.
    :type clear_line: bool
    :param bypass_silent_mode: If ``True``, prints the message even if silent mode is enabled.
    :type bypass_silent_mode: bool
    :return: ``True`` if the user chose "yes" for the yes/no prompt, ``False`` if they chose "no", or ``None`` if
        ``ask_yes_no`` is false.
    :rtype: bool | None
    """
    global silent_mode

    if silent_mode and not bypass_silent_mode:
        return True if ask_yes_no else None

    output: str = message
    if not ask_yes_no:
        output += end
    if clear_line:
        output = '\033[0K' + output

    if ask_yes_no:
        response: str = input(output + ' [y/n] \033[0m').strip().lower()
        return response in ('y', 'yes')
    else:
        sys.stdout.write(output)
        sys.stdout.flush()
        return None

def loading_message(message: str, clear_line: bool = True) -> None:
    """
    Prints a loading message and moves the cursor back to the start of the line.

    :param message: The message to print.
    :type message: str
    :param clear_line: Whether to clear all characters after the cursor on the current line before printing.
    :type clear_line: bool
    """
    generic_message(message, end='\r', clear_line=clear_line)

def success_message(message: str, clear_line: bool = True) -> None:
    """
    Prints a success message.

    :param message: The message to print.
    :type message: str
    :param clear_line: Whether to clear all characters after the cursor on the current line before printing.
    :type clear_line: bool
    """
    r: int; g: int; b: int
    r, g, b = SUCCESS_COLOR
    generic_message(
        f'\033[38;2;{r};{g};{b}m{message}\033[0m',
        clear_line=clear_line,
    )

def warning_message(message: str, ask_yes_no: bool = False, clear_line: bool = True) -> bool | None:
    """
    Prints a warning message.

    :param message: The message to print.
    :type message: str
    :param ask_yes_no: If ``True``, the user will be prompted for a yes/no answer (defaulting to "no" if invalid input,
        or "yes" if suppressed).
    :type ask_yes_no: bool
    :param clear_line: Whether to clear all characters after the cursor on the current line before printing.
    :type clear_line: bool
    :return: ``True`` if the user chose "yes" for the yes/no prompt, ``False`` if they chose "no", or ``None`` if
        ``ask_yes_no`` is false.
    :rtype: bool | None
    """
    global suppress_warnings

    if suppress_warnings:
        return True if ask_yes_no else None

    r: int; g: int; b: int
    r, g, b = WARNING_COLOR
    return generic_message(
        f'\033[38;2;{r};{g};{b}mWarning: {message}' + ('' if ask_yes_no else '\033[0m'),
        ask_yes_no=ask_yes_no,
        clear_line=clear_line,
        bypass_silent_mode=True,
    )

def exit_with_error(parser: argparse.ArgumentParser, message: str, status: int = 1, clear_line: bool = True) -> NoReturn:
    """
    Exits the program with an error message.

    :param parser: The argument parser.
    :type parser: argparse.ArgumentParser
    :param message: An error message to print.
    :type message: str
    :param status: The exit status code to exit with.
    :type status: int
    :param clear_line: Whether to clear all characters after the cursor on the current line before printing.
    :type clear_line: bool
    """
    r: int; g: int; b: int
    r, g, b = ERROR_COLOR
    generic_message(
        f'\033[38;2;{r};{g};{b}mError: {message}\033[0m',
        clear_line=clear_line,
        bypass_silent_mode=True,
    )

    parser.exit(status)

def int_in_range(min_value: int | None = None, max_value: int | None = None) -> Callable[[Any], int]:
    """
    Factory function that produces a callable that converts values into ``int`` instances, but raises ``ValueError`` if
    the resulting ``int`` would be outside a specified range, or if the input value cannot be converted.

    :param min_value: The minimum allowed return value of the callable, or ``None`` for no lower bound.
    :type min_value: int | None
    :param max_value: The maximum allowed return value of the callable, or ``None`` for no upper bound.
    :type max_value: int | None
    :return: A callable that converts values into ``int`` instances, or raises ``ValueError`` if they are out of bounds
        or cannot be converted.
    :rtype: Callable[[Any], int]
    """
    def convert(value: Any) -> int:
        nonlocal min_value
        nonlocal max_value

        # Convert to `int`
        int_: int = int(value)

        # Ensure in-bounds
        if min_value is not None and int_ < min_value:
            max_value_notation: int | str = '\u221e' if max_value is None else max_value
            raise ValueError(f'Value {int_} is outside the allowed range [{min_value}, {max_value_notation}]')
        if max_value is not None and int_ > max_value:
            min_value_notation: int | str = '-\u221e' if min_value is None else min_value
            raise ValueError(f'Value {int_} is outside the allowed range [{min_value_notation}, {max_value}]')

        return int_

    return convert

def hex_color(hex_code: str, is_rgba: bool = False) -> tuple[int, int, int] | tuple[int, int, int, int]:
    """
    Converts a hex color code with or without a leading ``#`` into an RGB or RGBA tuple.

    :param hex_code: The hex color code.
    :type hex_code: str
    :param is_rgba: If ``True``, alpha values other than ``255`` will be accepted, and outputs will be in RGBA format.
    :type is_rgba: bool
    :return: The color in RGB format if ``is_rgba`` is false, or RGBA format if ``is_rgba`` is true.
    :rtype: tuple[int, int, int] | tuple[int, int, int, int]
    :raises ValueError: If the hex color code is invalid.
    """
    hex_code = hex_code.removeprefix('#').lower()

    if len(hex_code) in (3, 4):
        hex_code = ''.join(digit * 2 for digit in hex_code)
    if len(hex_code) == 6:
        hex_code += 'ff'
    if len(hex_code) != 8:
        raise ValueError('Hex color code must have 3, 4, 6, or 8 digits')

    hex_parts: tuple[str, str, str, str] = (hex_code[:2], hex_code[2:4], hex_code[4:6], hex_code[6:])
    # noinspection PyTypeChecker
    rgba: tuple[int, int, int, int] = tuple(map(lambda hex_part: int(hex_part, 16), hex_parts))

    if rgba[3] != 255 and not is_rgba:
        raise ValueError('Transparency is not supported by this argument')

    return rgba if is_rgba else rgba[:3]

hex_color_rgb: Callable[[str], tuple[int, int, int]] = lambda hex_code: hex_color(hex_code, is_rgba=False)
hex_color_rgba: Callable[[str], tuple[int, int, int, int]] = lambda hex_code: hex_color(hex_code, is_rgba=True)

class NewlinePreservingHelpFormatter(argparse.HelpFormatter):
    """
    ``argparse`` help text formatter class that preserves all newline characters, indentation, and empty lines, while
    still wrapping individual lines.
    """

    def _split_lines(self, text: str, width: int) -> list[str]:
        # Split by newline characters
        explicit_lines: list[str] = text.split('\n')

        # Process each of these explicit lines individually
        result_lines: list[str] = []
        for explicit_line in explicit_lines:
            # Lines with only whitespace become empty lines
            if explicit_line.strip() == '':
                result_lines.append('')

            # Find indentation level of the line as a number of spaces
            indentation_level: int = 0
            for character in explicit_line:
                if character == ' ':
                    # Add 1 space
                    indentation_level += 1
                elif character == '\t':
                    # Snap to next multiple of 4 spaces
                    indentation_level = ((indentation_level + 4) // 4) * 4
                else:
                    break

            # Cap indentation level between 0 and 5 less than the width
            indentation_level = max(0, min(indentation_level, width - 5))

            # Wrap line
            wrapped_parts: list[str] = super()._split_lines(explicit_line, width - indentation_level)

            # Re-apply indentation
            wrapped_parts = [' ' * indentation_level + wrapped_part for wrapped_part in wrapped_parts]

            # Add to result
            result_lines += wrapped_parts

        return result_lines



# MAIN

def cli(args: list[str] | None = None) -> None:
    """
    Runs the WinFlame CLI (Command-Line Interface).

    :param args: A list of arguments passed to the CLI. If ``None``, the arguments passed to the current program are
        used.
    :type args: list[str] | None
    """
    global silent_mode
    global suppress_warnings

    # Compute paths

    program_dir: str = os.path.dirname(os.path.abspath(__file__))
    tree_file_hashes_path: str = os.path.join(program_dir, TREE_FILE_HASHES_FILENAME)

    cache_dir: str = os.path.join(program_dir, CACHE_DIR)
    cached_file_tree_path: str = os.path.join(cache_dir, CACHED_FILE_TREE_FILENAME)

    output_dir: str = os.path.join(program_dir, OUTPUT_DIR)


    # Argument parsing

    parser: argparse.ArgumentParser = argparse.ArgumentParser(
        prog = common.PROGRAM_NAME,
        description = common.PROGRAM_DESCRIPTION,
        add_help = False, # We manually add the help option back later so that we can customize it
        prefix_chars = '-/',
        formatter_class = NewlinePreservingHelpFormatter,
    )


    input_options_parent = parser.add_argument_group('Input options',
        description='Use one of these options to choose how to obtain a file tree.')
    input_options = input_options_parent.add_mutually_exclusive_group()

    input_options.add_argument('-b', '--build-from', '-t', '--target', metavar='ROOT',
        help='A file or directory to build the file tree from, called the "root" of the file tree.')
    input_options.add_argument('-i', '--tree-in', # -i for "in" or "import"
        help='A tree file to load that was created with -e. These files cannot be shared, as allowing this would allow'
            + ' arbitrary code execution via a deserialization attack. To prevent this, files loaded with this option'
            + ' are first checked against known hashes to verify that they were made on this device.')
    input_options.add_argument('-r', '--reuse-tree', action='store_true',
        help='Reuse the last file tree that was cached with -c.')


    input_config_options = parser.add_argument_group('Input configuration options',
        description='Extra configuration for the input options.')

    input_config_options.add_argument('-N', '--no-progress-report', action='store_true',
        help='Hide the progress report display when using -b.')


    output_options = parser.add_argument_group('Output options',
        description='Use these options to choose what to do with the file tree. You may use multiple at once.')

    flame_output_options = output_options.add_mutually_exclusive_group()
    flame_output_options.add_argument('-f', '-o', '--flame-out', # -f for "flame", -o for "out"
        help='Create a flame graph and write it to this file; supports all image formats that Pillow supports with'
            + ' RGBA.')
    flame_output_options.add_argument('-F', '-O', '--flame-out-default', action='store_true',
        help='Create a flame graph and write it to the program\'s output folder under a default name.')
    flame_output_options.add_argument('-p', '--preview-flame', action='store_true',
        help='Create a flame graph and open it in the default image program without saving it.')

    tree_output_options = output_options.add_mutually_exclusive_group()
    tree_output_options.add_argument('-e', '--tree-out', # -e for "export"
        help='Write the file tree to this file; can be loaded later with -i.')
    tree_output_options.add_argument('-E', '--tree-out-default', action='store_true',
        help='Write the file tree to the program\'s output folder under a default name.')

    output_options.add_argument('-c', '--cache-tree', action='store_true',
        help='Cache the file tree to be used again later with -r; overwrites the existing cached tree if there is one.')
    output_options.add_argument('-I', '--info', action='store_true',
        help='Print basic info about the file tree.')


    output_config_options = parser.add_argument_group('Output configuration options',
        description='Extra configuration for the output options.')

    output_config_options.add_argument('-V', '--open-flame', action='store_true', # -V for "view"
        help='Open the flame graph in the default image program once it is completed when using -f or -F (implied when'
            + ' using -p).')


    flame_graph_options = parser.add_argument_group('Flame graph options',
        description='Configuration for the flame graph when using -f, -F, or -p.')

    flame_graph_options.add_argument('-R', '--flame-root',
        help='Build the flame graph from a different node of the file tree than the root. If provided, this should be a'
            + ' path relative to the file tree\'s root that does not traverse any symlinks or junctions. If you don\'t'
            + ' know the file tree\'s root, you can check with -I. In the special case where the file tree root is a'
            + ' file and you want to make the flame root one of its alternate data streams (for some reason), just the'
            + ' stream suffix should be used (e.g. \':ads\' or \':ads:$DATA\').')
    flame_graph_options.add_argument('-m', '--max-flame-depth', type=int_in_range(0, None),
        help='Limit the number of layers above the flame root to draw.')

    flame_graph_options.add_argument('-L', '--logical-size', action='store_true',
        help='Use files\' logical size instead of their physical size for proportions.')

    flame_graph_options.add_argument('-w', '--width', type=int_in_range(2, None), default=1920,
        help='Width of the graph in pixels. (Default: 1920)')
    flame_graph_options.add_argument('-H', '--layer-height', type=int_in_range(2, None), default=20,
        help='Height of each graph layer in pixels. (Default: 20)')

    flame_graph_options.add_argument('-B', '--bg-color', type=hex_color_rgba, default=(255, 255, 255, 255),
        help='Color of the graph\'s background as a hex color code with an optional \'#\' prefix. (Default: #ffff)')
    flame_graph_options.add_argument('-T', '--fg-color', type=hex_color_rgba, default=(0, 0, 0, 255), # -T for "text"
        help='Color of rectangle outlines and labels as a hex color code with an optional \'#\' prefix. (Default:'
            + ' #000f)')

    flame_graph_options.add_argument('-l', '--labels', choices=['none', 'files', 'special', 'all'], default='all',
        help='''Visibility of labels. (Default: 'all')

Allowed values:
    'none': All labels are hidden.
    'files': Labels on file tree nodes are shown.
    'special': Labels on special segments are shown. See help on -S for more info.
    'all': All labels are shown.
''')
    flame_graph_options.add_argument('-W', '--min-label-width', type=int_in_range(1, None), default=15,
        help='Minimum width of a rectangle in pixels for a label to be drawn on it. (Default: 15)')

    flame_graph_options.add_argument('-g', '--font-file', # -g because I ran out of letters
        help='Font file to use for labels. Accepts TTF, OTF, PCF, BDF, and PIL font files. If omitted, a default font'
            + ' is used.')
    # In the below help text, the "%" symbol is doubled to escape it because argument help texts are formatted strings.
    flame_graph_options.add_argument('-G', '--font-size', type=int_in_range(1, None), # -G to match -g for --font-file
        help='Font size to use for labels in pixels. This only has an effect if a TTF or OTF font file is used. If'
            + ' omitted, 60%% of the value of -H is used, with a cap of 10.')

    flame_graph_options.add_argument('-S', '--special', choices=['none', 'unaccounted', 'unaccounted-free', 'all'], default='all',
        help='''Visibility of special segments. (Default: 'all')

Allowed values:
    'none': All special segments are hidden.
    'unaccounted': The Unaccounted special segment is shown if available.
    'unaccounted-free': The Unaccounted and Free special segments are shown if available.
    'all': All special segments are shown if available.

Special segments:
    Special segments are extra rectangles that get drawn on the flame graph at the bottom layer to display'''
    + ''' information about the drive, and they are only available if the root of the graph is a drive root.'''
    + ''' Additionally, special segments are always hidden if -L is supplied because they have no logical size'''
    + ''' equivalents.

    Free: Represents the amount of free (unused) space on the drive.
    Unaccounted*: Represents the amount of used space on the drive that the program could not identify the source of.
    Extra-counted*: Represents the amount of unused space on the drive that the program actually over-counted as used'''
    + ''' by files.
    
    *This special segment is not available unless the file tree was built with administrator privileges due to'''
    + ''' Windows API limitations.''')


    miscellaneous_options = parser.add_argument_group('Miscellaneous options')

    miscellaneous_action_options = miscellaneous_options.add_mutually_exclusive_group()
    miscellaneous_action_options.add_argument('-h', '--help', '/?', action='store_true',
        help='Print this help text and exit.')
    miscellaneous_action_options.add_argument('-v', '--version', action='store_true',
        help='Print program version information and exit.')
    miscellaneous_action_options.add_argument('-C', '--colors', action='store_true',
        help='Print the flame graph / progress report color key and exit.')
    miscellaneous_action_options.add_argument('-d', '--delete-cache', '--clear-cache', action='store_true', # -d for "delete"
        help='Delete the file tree cached with -c (if there is one) and exit.')

    miscellaneous_options.add_argument('-s', '--silent', action='store_true',
        help='Suppress all output except errors and warnings.')
    miscellaneous_options.add_argument('-n', '--no-warnings', action='store_true',
        help='Suppress all warnings.')


    args: argparse.Namespace = parser.parse_args(args)


    # Miscellaneous options

    # Output suppression
    silent_mode = args.silent
    suppress_warnings = args.no_warnings

    # Help text
    if args.help:
        if args.silent:
            exit_with_error(parser, 'Cannot show help with --silent.')

        # Print help and exit
        parser.print_help()
        parser.exit()

    # Version information
    if args.version:
        if args.silent:
            exit_with_error(parser, 'Cannot show version information with --silent.')

        # Print version information and exit
        python_revision: str = platform.python_revision()
        generic_message(f'{common.PROGRAM_NAME} {common.VERSION} running on {platform.platform()} with'
            + f' {platform.python_implementation()} {platform.python_version()}'
            + (f' (revision {python_revision})' if python_revision != '' else ''))
        parser.exit()

    # Color key
    if args.colors:
        if args.silent:
            exit_with_error(parser, 'Cannot show color key with --silent.')

        output: str = ''

        for line in COLOR_KEY:
            # Just add a newline if this is a blank line
            if line is None:
                output += '\n'
                continue

            # Get line color
            r: int; g: int; b: int
            r, g, b = line[1]

            # Set terminal text color
            output += f'\033[38;2;{r};{g};{b}m'

            # Write color hex code
            output += f'#{hex(r)[2:].zfill(2)}{hex(g)[2:].zfill(2)}{hex(b)[2:].zfill(2)} '

            # Write color label
            output += line[0]

            # Reset terminal formatting and add newline
            output += '\033[0m\n'

        # Footer
        output += '\n*Visible on flame graphs.\n'

        # Print color key and exit
        sys.stdout.write(output)
        sys.stdout.flush()
        parser.exit()

    # Clear cache
    if args.delete_cache:
        if os.path.exists(cache_dir):
            loading_message('Clearing cache...')
            shutil.rmtree(cache_dir)
            success_message('Cache cleared.')
        else:
            success_message('The cache is already empty.')

        parser.exit()


    # Determine I/O modes

    # Input
    input_mode_build_tree: bool = args.build_from is not None
    input_mode_tree_file: bool = args.tree_in is not None
    input_mode_reuse_tree: bool = args.reuse_tree

    # Output
    output_mode_flame: bool = args.flame_out is not None or args.flame_out_default or args.preview_flame
    output_mode_tree_file: bool = args.tree_out is not None or args.tree_out_default
    output_mode_cache_tree: bool = args.cache_tree
    output_mode_tree_info: bool = args.info

    # Check if at least one option is present for each category
    input_option_is_present: bool = (
            input_mode_build_tree
            or input_mode_tree_file
            or input_mode_reuse_tree
    )
    output_option_is_present: bool = (
            output_mode_flame
            or output_mode_tree_file
            or output_mode_cache_tree
            or output_mode_tree_info
    )

    # Ensure that an input option is present if an output option is present and vice versa
    if output_option_is_present and not input_option_is_present:
        exit_with_error(parser, 'Output(s) specified but no input provided.')
    elif input_option_is_present and not output_option_is_present:
        exit_with_error(parser, 'Input provided but no output(s) specified.')

    # Error if both --info and -silent are passed
    if output_mode_tree_info and args.silent:
        exit_with_error(parser, 'Cannot show file tree info with --silent.')


    # Obtain file tree

    file_tree_root: winflame.FileNode

    if input_mode_build_tree:
        # Build a file tree from a target (root) path

        # Normalize path
        normalized_target_path: str = os.path.normpath(args.build_from)

        # Ensure path exists
        if not os.path.exists(normalized_target_path):
            exit_with_error(parser, f'Failed to build file tree from {normalized_target_path!r}; target path does not exist.')

        # Absolutize path and ensure ":$DATA" suffix if it is an alternate data stream
        root_path: str = os.path.abspath(normalized_target_path)
        filename: str = os.path.basename(root_path)
        filename_colon_count: int = filename.count(':')
        if filename_colon_count == 1:
            # Add ":$DATA" suffix
            filename += ':$DATA'
            root_path = os.path.join(os.path.dirname(root_path), filename)
        elif filename_colon_count > 1:
            if filename_colon_count != 2 or not filename.endswith(':$DATA'):
                exit_with_error(parser, f'Failed to build file tree from {normalized_target_path!r}; invalid alternate data stream syntax.')

        # Show warning and ask for confirmation if building from a drive root without elevated privileges
        canonical_root_path: str = os.path.realpath(root_path, strict=True)
        drive_letter: str; path_on_drive: str
        drive_letter, path_on_drive = os.path.splitdrive(canonical_root_path)
        is_drive_root: bool = drive_letter != '' and path_on_drive == os.path.sep
        if is_drive_root and not win32com.shell.shell.IsUserAnAdmin():
            should_continue: bool = warning_message('You are trying to build a file tree from a drive root without'
                + ' administrator privileges. Not all information may be available, for example drive capacity and'
                + ' system file sizes. Continue anyway?', ask_yes_no=True)
            if not should_continue:
                parser.exit()

        # Build file tree
        loading_message('Building file tree...')
        file_tree_root = winflame.FileNode(
            root_path,
            is_ads = root_path.endswith(':$DATA'),
            should_report_progress = not (args.no_progress_report or args.silent),
        )
        success_message(f'Built file tree from {root_path!r}.')

    elif input_mode_tree_file:
        # Load a file tree from a file

        # Normalize path
        tree_file_path: str = os.path.normpath(args.tree_in)

        # Ensure path exists and is a file
        if not os.path.isfile(tree_file_path):
            exit_with_error(parser, f'Failed to load file tree {tree_file_path!r}; file does not exist.')

        # Verify and load file tree
        loading_message('Loading file tree...')
        load_result: winflame.FileNode | None = load_file_tree(tree_file_path, tree_file_hashes_path)
        if load_result is None:
            exit_with_error(parser, f'Failed to load file tree {tree_file_path!r}; file did not match any known hashes. Was it created by this program installation?')

        # Use tree if it was verified as safe
        file_tree_root = load_result
        success_message(f'Loaded file tree from {tree_file_path!r}.')

    elif input_mode_reuse_tree:
        # Load cached file tree

        # Ensure cached file tree exists
        if not os.path.isfile(cached_file_tree_path):
            exit_with_error(parser, 'There is no cached file tree.')

        # Verify and load cached file tree
        loading_message('Loading cached file tree...')
        load_result: winflame.FileNode | None = load_file_tree(cached_file_tree_path, tree_file_hashes_path)
        if load_result is None:
            exit_with_error(parser, 'The cached file tree did not match any known hashes.')

        # Use tree if it was verified as safe
        file_tree_root = load_result
        success_message('Loaded cached file tree.')


    # Output

    if output_mode_tree_info:
        # Print basic info about the file tree

        # noinspection PyUnboundLocalVariable
        generic_message(f'''File tree info:
    Root path: {file_tree_root.path!r}
    Root file type: {file_tree_root.file_type.human_readable_name}
    Total physical size: {winflame.format_data_size(file_tree_root.total_physical_size)} ({file_tree_root.total_physical_size:,} bytes)
    Total logical size: {winflame.format_data_size(file_tree_root.total_logical_size)} ({file_tree_root.total_logical_size:,} bytes)
    Is drive root: {str(file_tree_root.is_drive_root).lower()}
    Drive capacity: {
        (
            'Unknown (insufficient permissions)'
                if file_tree_root.drive_capacity is None
                    else f'{winflame.format_data_size(file_tree_root.drive_capacity)} ({file_tree_root.drive_capacity:,} bytes)'
        )
            if file_tree_root.is_drive_root
                else 'N/A'
    }
    Drive free space: {
        f'{winflame.format_data_size(file_tree_root.drive_free_space)} ({file_tree_root.drive_free_space:,} bytes)'
            if file_tree_root.is_drive_root
                else 'N/A'
    }
    Build timestamp: {time.strftime('%m/%d/%Y, %I:%M:%S %p', time.localtime(file_tree_root.build_timestamp))}''')

        # Count nodes
        loading_message('    Node count: loading...')
        node_count: int = 0
        error_node_count: int = 0
        for node in file_tree_root.descendants_iter(include_self=True):
            node_count += 1
            if node.is_error:
                error_node_count += 1
        generic_message(f'    Node count: {node_count:,} (including {error_node_count:,} error{"" if error_node_count == 1 else "s"})')

    if output_mode_cache_tree:
        # Save file tree to the cache

        # Create cache folder if it doesn't exist
        if not os.path.exists(cache_dir):
            os.mkdir(cache_dir)

        # Save to cache
        loading_message('Caching file tree...')
        save_file_tree(file_tree_root, cached_file_tree_path, tree_file_hashes_path)
        cached_file_tree_size: int = os.path.getsize(cached_file_tree_path)
        success_message(f'Cached file tree ({winflame.format_data_size(cached_file_tree_size)}).')

    if output_mode_tree_file:
        # Export file tree to a file

        # Choose output file path
        file_tree_path: str
        if args.tree_out_default:
            # Generate default filename from current time
            file_tree_filename: str = time.strftime(DEFAULT_FILE_TREE_FILENAME_FORMAT)
            file_tree_path = os.path.join(output_dir, file_tree_filename)

            # Create output folder if it doesn't exist
            if not os.path.exists(output_dir):
                os.mkdir(output_dir)
        else:
            file_tree_path = args.tree_out

        # Ensure path isn't reserved
        if os.path.isreserved(file_tree_path):
            exit_with_error(parser, f'Cannot export file tree to reserved path {file_tree_path!r}. Does it contain a reserved character?')

        # Export
        loading_message('Exporting file tree...')
        save_file_tree(file_tree_root, file_tree_path, tree_file_hashes_path)
        exported_file_tree_size: int = os.path.getsize(file_tree_path)
        success_message(f'Wrote file tree to {os.path.relpath(file_tree_path)!r} ({winflame.format_data_size(exported_file_tree_size)}).')

    if output_mode_flame:
        # Create a flame graph

        # Choose output file path
        flame_graph_path: str | None # `None` indicates preview
        if args.flame_out_default:
            # Generate default filename from current time
            flame_graph_filename: str = time.strftime(DEFAULT_FLAME_GRAPH_FILENAME_FORMAT)
            flame_graph_path = os.path.join(output_dir, flame_graph_filename)

            # Create output folder if it doesn't exist
            if not os.path.exists(output_dir):
                os.mkdir(output_dir)
        elif args.preview_flame:
            flame_graph_path = None
        else:
            flame_graph_path = args.flame_out

        # Ensure output path isn't reserved
        if flame_graph_path is not None and os.path.isreserved(flame_graph_path):
            exit_with_error(parser, f'Cannot write flame graph to reserved path {flame_graph_path!r}. Does it contain a reserved character?')

        # Locate flame graph root node by gradually moving it down the tree as needed
        flame_root: winflame.FileNode = file_tree_root
        if args.flame_root is not None:
            # Ensure path is relative
            if os.path.isabs(args.flame_root):
                exit_with_error(parser, f'Flame root path {args.flame_root!r} is absolute; it must be relative to the file tree root.')

            # Ensure path does not back out
            normalized_flame_root_path: str = os.path.normpath(args.flame_root)
            if normalized_flame_root_path.startswith('..'):
                exit_with_error(parser, f'Flame root path {args.flame_root!r} backs out too far; it must be on the file tree.')

            # Traverse path
            for path_segment in normalized_flame_root_path.split(os.path.sep):
                # Do not move down the tree if this path segment simply references ourselves
                if path_segment == '.':
                    continue

                # Get node children and a mapping with its child names all lowercase
                children: dict[str, winflame.FileNode] = flame_root.children
                children_lower: dict[str, winflame.FileNode] = {
                    child_name.lower(): child for child_name, child in children.items()
                }

                # Parse stream suffix if the path segment contains a colon (which indicates an alternate data stream)
                stream_suffix: str | None = None
                if ':' in path_segment:
                    # Ensure ":$DATA" suffix
                    path_segment = path_segment.removesuffix(':$DATA') + ':$DATA'

                    # Separate filename from stream suffix
                    partition_result: tuple[str, str, str] = path_segment.partition(':')
                    filename: str
                    filename, stream_suffix = partition_result[0], partition_result[1] + partition_result[2]

                    # This is not an alternate data stream if the stream name is blank
                    if stream_suffix == '::$DATA':
                        stream_suffix = None

                    # Handle the filename as its own path segment
                    path_segment = filename

                # Do not move down the file tree if this is an alternate data stream and the filename is blank
                if not (path_segment == '' and stream_suffix is not None):
                    # Move down the file tree
                    if path_segment in children:
                        # Case-sensitive check
                        flame_root = children[path_segment]
                    elif path_segment.lower() in children_lower:
                        # Case-insensitive check
                        flame_root = children_lower[path_segment.lower()]
                    else:
                        exit_with_error(parser, f'Couldn\'t find flame root {args.flame_root!r} on file tree (first missing path segment is {path_segment!r}).')

                # Finish handling alternate data stream
                if stream_suffix is not None:
                    # Get new node children and a mapping with its child names all lowercase
                    children = flame_root.children
                    children_lower = {child_name.lower(): child for child_name, child in children.items()}

                    # Move down the file tree again, this time to the alternate data stream
                    if stream_suffix in children:
                        # Case-sensitive check
                        flame_root = children[stream_suffix]
                    elif stream_suffix.lower() in children_lower:
                        # Case-insensitive check
                        flame_root = children_lower[stream_suffix.lower()]
                    else:
                        exit_with_error(parser, f'Couldn\'t find flame root {args.flame_root!r} on file tree (couldn\'t find alternate data stream {stream_suffix!r} on file {path_segment!r}).')

        # Compute font size for labels
        font_size: float
        if args.font_size is None:
            font_size = min(0.6 * args.layer_height, 10.0)
        else:
            font_size = args.font_size

        # Load font for labels
        loading_message('Loading font...')
        font: ImageFont.ImageFont | ImageFont.FreeTypeFont
        if args.font_file is None:
            # Default font
            font = ImageFont.truetype('C:\\Windows\\Fonts\\bahnschrift.ttf', size=font_size)
        else:
            # Ensure file exists
            if not os.path.isfile(args.font_file):
                exit_with_error(parser, f'Failed to load font {args.font_file!r}; file does not exist.')

            # Get file extension
            font_file_extension: str = os.path.splitext(args.font_file)[1].lower()

            # Load font file
            match font_file_extension:
                case '.ttf' | '.otf':
                    font = ImageFont.truetype(args.font_file, size=font_size)
                case '.pcf':
                    with open(args.font_file, 'rb') as f:
                        font = PcfFontFile.PcfFontFile(f).to_imagefont()
                case '.bdf':
                    with open(args.font_file, 'rb') as f:
                        font = BdfFontFile.BdfFontFile(f).to_imagefont()
                case _: # Assumed to be PIL format
                    font = ImageFont.load(args.font_file)

        # Create flame graph
        loading_message('Creating flame graph...')
        # noinspection PyUnboundLocalVariable
        flame_graph: Image.Image = flame_root.create_flame_graph(
            use_physical_size = not args.logical_size,
            max_depth = args.max_flame_depth,
            width = args.width,
            layer_height = args.layer_height,
            background_color = args.bg_color,
            foreground_color = args.fg_color,
            label_visibility = winflame.LabelVisibility(args.labels),
            min_label_width = args.min_label_width,
            font = font,
            show_drive_free_space = args.special in ('unaccounted-free', 'all'),
            show_drive_unaccounted_space = args.special in ('unaccounted', 'unaccounted-free', 'all'),
            show_drive_extra_counted_space = args.special == 'all',
        )

        # Save or preview flame graph
        if flame_graph_path is not None:
            # Write to file
            loading_message('Saving flame graph...')
            flame_graph.save(flame_graph_path)
            success_message(f'Wrote flame graph to {os.path.relpath(flame_graph_path)!r}.')

            # Open in default image program if flag is enabled
            if args.open_flame:
                os.startfile(flame_graph_path)
        else:
            # Preview
            loading_message('Opening flame graph...')
            flame_graph.show()
            success_message('Opened flame graph for previewing.')


    # Print usage if no I/O options were given and we did not already exit from completing a miscellaneous action

    if not (input_option_is_present or output_option_is_present or args.silent):
        parser.print_usage()
        generic_message(f'\nType {parser.prog} --help for detailed help.')


if __name__ == '__main__':
    cli()
