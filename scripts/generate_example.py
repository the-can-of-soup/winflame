"""
Helper script that generates an example file tree to showcase WinFlame with.
"""

# IMPORTS

from __future__ import annotations

from collections.abc import Callable
import string
import random
import shutil
import math
import sys
import os



# CONSTANTS

MIN_DIRECTORY_CHILDREN: int = 1 # Minimum number of children (files/directories) that will be put into each directory
MAX_DIRECTORY_CHILDREN_LEVELS: list[int] = [5, 5, 4, 4, 3, 3, 2, 2, 1] # Maximum of the above metric at each depth level, with larger depths using the final value of the list
DIRECTORY_CHANCE_FUNC: Callable[[int], float] = lambda depth: max(0.1, 0.9 - 0.025 * depth) # Function that returns the probability that a child of a directory is a directory (otherwise it is a file) given the depth of the parent
MIN_FILE_SIZE_FUNC: Callable[[int], int] = lambda depth: math.ceil(5_120 / (depth + 1)) # Function that returns the minimum logical size of a file in bytes at a given depth
MAX_FILE_SIZE_FUNC: Callable[[int], int] = lambda depth: math.ceil(10_240 / (depth + 1)) # Same as above, but for maximum
ALTERNATE_DATA_STREAM_CHANCE: float = 0.2 # Probability that a file is given alternate data streams
MIN_ALTERNATE_DATA_STREAM_COUNT: int = 1 # Minimum number of alternate data streams that a file may be given
MAX_ALTERNATE_DATA_STREAM_COUNT: int = 2 # Maximum of the above metric
MIN_ALTERNATE_DATA_STREAM_SIZE_FUNC: Callable[[int], int] = lambda depth: math.ceil(2_048 / (depth + 1)) # Function that returns the minimum logical size of an alternate data stream in bytes at a given depth
MAX_ALTERNATE_DATA_STREAM_SIZE_FUNC: Callable[[int], int] = lambda depth: math.ceil(4_096 / (depth + 1)) # Same as above, but for maximum

LIPSUM: str = ('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et'
    + ' dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea'
    + ' commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla'
    + ' pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est'
    + ' laborum.')

# Relative to directory containing this file
EXAMPLE_DIR: str = '../example'



# GLOBALS

# Get lipsum words
lipsum_words: list[str] = LIPSUM.split(' ')
for punctuation_character in string.punctuation:
    lipsum_words = [word.replace(punctuation_character, '') for word in lipsum_words]

# Index into `lipsum_words` to pull the next word from
lipsum_index: int = 0

# Get lipsum bytes
#
# A space is appended because this bytestring will be repeated.
lipsum_bytes: bytes = (LIPSUM + ' ').encode()



# DEFINITIONS

def get_next_unique_lipsum_words(word_count: int) -> list[str]:
    """
    Gets some unique lipsum words.

    The returned words are unique to each other, but have no guarantees about uniqueness with other words returned by
    previous calls of this function.

    :param word_count: The number of unique words to get.
    :type word_count: int
    :return: A list of unique lipsum words.
    :rtype: list[str]
    """
    global lipsum_words
    global lipsum_index

    words: list[str] = []
    for _ in range(word_count):
        base_word: str = lipsum_words[lipsum_index]

        # Append first unique number suffix
        word_candidate: str = base_word
        candidate_number: int = 1
        while word_candidate in words:
            candidate_number += 1
            word_candidate = base_word + str(candidate_number)

        words.append(word_candidate)
        lipsum_index = (lipsum_index + 1) % len(lipsum_words)

    return words

def create_example_file_tree_at(
        path: str,
        current_depth: int,
        is_directory: bool = True,
        is_ads: bool = False,
    ) -> None:
    """
    Recursively creates an example file tree at the specified path.

    :param path: The path to create the file tree at.
    :type path: str
    :param current_depth: The current depth within the file tree starting from zero.
    :type current_depth: int
    :param is_directory: ``True`` if ``path`` refers to a directory rather than a file or alternate data stream.
    :type is_directory: bool
    :param is_ads: ``True`` if ``path`` refers to an alternate data stream.
    :type is_ads: bool
    """
    global lipsum_words
    global lipsum_bytes

    if is_directory:
        # Create directory
        os.mkdir(path)

        # Choose child count
        max_children: int = MAX_DIRECTORY_CHILDREN_LEVELS[current_depth if current_depth < len(MAX_DIRECTORY_CHILDREN_LEVELS) else -1]
        child_count: int = random.randint(MIN_DIRECTORY_CHILDREN, max_children)

        # Choose child lipsum words
        child_words: list[str] = get_next_unique_lipsum_words(child_count)

        # Recursively create children
        for child_word in child_words:
            child_is_directory: bool = random.random() < DIRECTORY_CHANCE_FUNC(current_depth)
            create_example_file_tree_at(
                os.path.join(path, child_word + ('' if child_is_directory else '.txt')),
                current_depth + 1,
                is_directory = child_is_directory,
                is_ads = False,
            )

    else:
        # Choose file / alternate data stream size
        file_size: int
        if is_ads:
            file_size = random.randint(
                MIN_ALTERNATE_DATA_STREAM_SIZE_FUNC(current_depth),
                MAX_ALTERNATE_DATA_STREAM_SIZE_FUNC(current_depth),
            )
        else:
            file_size = random.randint(
                MIN_FILE_SIZE_FUNC(current_depth),
                MAX_FILE_SIZE_FUNC(current_depth),
            )

        # Repeat the lipsum text as the data
        file_data: bytes = b''
        file_data += lipsum_bytes * (file_size // len(lipsum_bytes))
        file_data += lipsum_bytes[:file_size % len(lipsum_bytes)]

        # Create file / alternate data stream
        with open(path, 'wb') as f:
            f.write(file_data)

        # If this is an alternate data stream, we are done
        if is_ads:
            return

        # Choose alternate data stream count
        ads_count: int = 0
        if random.random() < ALTERNATE_DATA_STREAM_CHANCE:
            ads_count = random.randint(MIN_ALTERNATE_DATA_STREAM_COUNT, MAX_ALTERNATE_DATA_STREAM_COUNT)

        # Choose ADS lipsum words
        ads_words: list[str] = get_next_unique_lipsum_words(ads_count)

        # Recursively create alternate data streams
        for ads_word in ads_words:
            create_example_file_tree_at(
                f'{path}:{ads_word}:$DATA',
                current_depth + 1,
                is_directory = False,
                is_ads = True,
            )



# MAIN

if __name__ == '__main__':
    # Change to directory containing this file
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Delete old example (if there is one)
    if os.path.isdir(EXAMPLE_DIR):
        was_confirmed: bool = input(f'Delete the old example file tree at {EXAMPLE_DIR!r}? [y/n] ').strip().lower() in ('y', 'yes')
        if was_confirmed:
            shutil.rmtree(EXAMPLE_DIR)
        else:
            sys.exit()

    # Create example file tree
    print('Creating example file tree...')
    create_example_file_tree_at(EXAMPLE_DIR, 0)
    print(f'Done! Example file tree created at {EXAMPLE_DIR!r}.')
