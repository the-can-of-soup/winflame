# `pywin32` documentation: https://mhammond.github.io/pywin32/
# Windows API reference: https://learn.microsoft.com/en-us/windows/win32/api/
# Windows Data Access and Storage API documentation: https://learn.microsoft.com/en-us/windows/win32/api/_fs/
# Alternate Windows API list (not sure what these are exactly): https://learn.microsoft.com/en-us/windows/win32/apiindex/windows-api-list



# IMPORTS

from __future__ import annotations

# Validate environment
import platform
if platform.system() != 'Windows':
    raise RuntimeError('This program only supports Windows')

import win32file, winioctlcon, win32api, win32com.shell.shell, pywintypes
from PIL import Image, ImageDraw, ImageFont
from collections.abc import Iterator
from collections import OrderedDict
import enum
import time
import math
import sys
import os



# DEFINITIONS

def format_data_size(size_in_bytes: int, use_iec_units: bool = True, max_fractional_digits: int = 2) -> str:
    """
    Formats a data size in bytes into a short human-readable string.

    :param size_in_bytes: The data size in bytes.
    :type size_in_bytes: int
    :param use_iec_units: If ``True``, the IEC units are used (kibibyte, mebibyte, gibibyte, etc.). If ``False``, the SI
        units are used (kilobyte, megabyte, gigabyte, etc.).
    :type use_iec_units: bool
    :param max_fractional_digits: The maximum number of fractional digits (digits after the decimal point) to use.
    :type max_fractional_digits: int
    :return: The formatted data size.
    :rtype: str
    """
    # Choose unit names
    unit_names: list[str] = [
        'KiB', 'MiB', 'GiB', 'TiB', 'PiB', 'EiB', 'ZiB', 'YiB', 'RiB', 'QiB',
    ] if use_iec_units else [
        'kB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB', 'RB', 'QB',
    ]

    # Choose unit
    unit_value: int = 1
    unit_name: str = 'B'
    for new_unit_name in unit_names:
        new_unit_value: int = unit_value * (1_024 if use_iec_units else 1_000)

        if new_unit_value > size_in_bytes:
            break
        unit_value = new_unit_value
        unit_name = new_unit_name

    # Format
    size_in_unit_rounded: float = round(size_in_bytes / unit_value, ndigits=max_fractional_digits)
    if size_in_unit_rounded.is_integer():
        size_in_unit_rounded = int(size_in_unit_rounded)
    return f'{size_in_unit_rounded:,} {unit_name}'


def truncate(text: str, width: int) -> str:
    """
    Truncates a string if necessary to stay within a character limit.

    :param text: The string to truncate.
    :type text: str
    :param width: The character limit.
    :type width: int
    :return: The truncated string, or the string as-is if it does not exceed the character limit.
    :rtype: str
    """
    if len(text) > width:
        return text[:width - 1] + '\u2026' # Truncate and replace last character with ellipsis
    return text


class LabelVisibility(enum.Enum):
    """
    A visibility setting for labels on a flame graph.

    See also: ``FileNode.create_flame_graph``
    """

    NONE = 'none'
    FILES = 'files'
    SPECIAL_SEGMENTS = 'special'
    ALL = 'all'

    def is_visible(self, is_node: bool) -> bool:
        """
        Checks if a label should be drawn on a flame graph with this visibility setting.

        :param is_node: ``True`` if the label is for a node rather than a special segment.
        :type is_node: bool
        :return: ``True`` if the label should be drawn.
        :rtype: bool
        """
        match self:
            case LabelVisibility.NONE: return False
            case LabelVisibility.FILES: return is_node
            case LabelVisibility.SPECIAL_SEGMENTS: return not is_node
            case LabelVisibility.ALL: return True


class WindowsFileAttributes(enum.Flag):
    """
    A set of a Windows file attributes.

    Source: https://learn.microsoft.com/en-us/windows/win32/fileio/file-attribute-constants
    """

    FILE_ATTRIBUTE_READONLY = 1
    FILE_ATTRIBUTE_HIDDEN = 2
    FILE_ATTRIBUTE_SYSTEM = 4
    FILE_ATTRIBUTE_DIRECTORY = 16
    FILE_ATTRIBUTE_ARCHIVE = 32
    FILE_ATTRIBUTE_DEVICE = 64 # Windows internal use only
    FILE_ATTRIBUTE_NORMAL = 128
    FILE_ATTRIBUTE_TEMPORARY = 256
    FILE_ATTRIBUTE_SPARSE_FILE = 512
    FILE_ATTRIBUTE_REPARSE_POINT = 1024
    FILE_ATTRIBUTE_COMPRESSED = 2048
    FILE_ATTRIBUTE_OFFLINE = 4096
    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 8192
    FILE_ATTRIBUTE_ENCRYPTED = 16384
    FILE_ATTRIBUTE_INTEGRITY_STREAM = 32768
    FILE_ATTRIBUTE_VIRTUAL = 65536 # Windows internal use only
    FILE_ATTRIBUTE_NO_SCRUB_DATA = 131072
    FILE_ATTRIBUTE_RECALL_ON_OPEN = 262144
    FILE_ATTRIBUTE_EA = 262144 # Windows internal use only
    FILE_ATTRIBUTE_PINNED = 524288
    FILE_ATTRIBUTE_UNPINNED = 1048576
    FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 4194304


class FileType(enum.Enum):
    """
    A type of ``FileNode``.

    The type ``FileType.HARD_LINK`` is assigned to nodes that have multiple names, except the first name discovered. For
    example, if the names ``C:\\foo``, ``C:\\bar``, and ``C:\\baz`` point to the same regular file, then the first name
    that is discovered will have the type ``FileType.REGULAR_FILE``, and the other two will have ``FileType.HARD_LINK``.
    """
    UNKNOWN = -1
    DIRECTORY = 0
    REGULAR_FILE = 1
    ALTERNATE_DATA_STREAM = 2
    SYMLINK = 3
    JUNCTION = 4
    HARD_LINK = 5

    @property
    def can_store_data(self) -> bool:
        """
        ``True`` if files of this type can intrinsically store data.

        Only true for ``FileType.REGULAR_FILE`` and ``FileType.ALTERNATE_DATA_STREAM``.

        See also: ``FileNode.can_store_data``

        :rtype: bool
        """
        return self in (FileType.REGULAR_FILE, FileType.ALTERNATE_DATA_STREAM)

    @property
    def human_readable_name(self) -> str:
        """
        A lowercase human-readable name for the file type.

        See also: ``FileType.short_human_readable_name``

        :rtype: str
        """
        return {
            FileType.UNKNOWN: 'unknown type',
            FileType.DIRECTORY: 'directory',
            FileType.REGULAR_FILE: 'regular file',
            FileType.ALTERNATE_DATA_STREAM: 'alternate data stream',
            FileType.SYMLINK: 'symlink',
            FileType.JUNCTION: 'junction',
            FileType.HARD_LINK: 'hard link',
        }[self]

    @property
    def short_human_readable_name(self) -> str:
        """
        A short, lowercase, human-readable name for the file type.

        See also: ``FileType.human_readable_name``

        :rtype: str
        """
        return {
            FileType.UNKNOWN: 'unknown type',
            FileType.DIRECTORY: 'directory',
            FileType.REGULAR_FILE: 'file',
            FileType.ALTERNATE_DATA_STREAM: 'alt. data stream',
            FileType.SYMLINK: 'symlink',
            FileType.JUNCTION: 'junction',
            FileType.HARD_LINK: 'hard link',
        }[self]

    @property
    def color_rgb(self) -> tuple[int, int, int]:
        """
        A color in RGB format to display the file type.

        :rtype: tuple[int, int, int]
        """
        return FILE_TYPE_COLORS[self]


# noinspection PyProtectedMember
class FileNode:
    def __init__(
            self,
            filename_or_path: str,
            parent: FileNode | None = None,
            is_ads: bool = False,
            catch_errors: bool = True,
            should_report_progress: bool = True,
    ) -> None:
        """
        A node of a file tree.

        When ``__init__`` is called, the filesystem will be scanned recursively to build the node's descendants.

        :param filename_or_path: The filename of the node excluding parent directories, unless this is a root node, in
            which case it should be the absolute path of the node.
        :type filename_or_path: str
        :param parent: The parent node, or ``None`` if this is a root node.
        :type parent: FileNode | None
        :param is_ads: ``True`` if this is an alternate data stream.
        :type is_ads: bool
        :param catch_errors: If ``True``, filesystem-related errors will be caught and saved to ``self.error``. If this
            happens, ``self.file_type`` will be set to ``FileType.UNKNOWN``, all file sizes will be set to zero, and
            some other attributes will be set to their empty state.
        :type catch_errors: bool
        :param should_report_progress: If ``True``, a formatted progress report will be printed to standard output
            periodically until initialization is complete. This attribute has no effect on non-root nodes.
        :type should_report_progress: bool
        """

        # Set basic attributes

        assert (parent is None) is os.path.isabs(filename_or_path)
        if parent is not None:
            assert os.path.dirname(filename_or_path) == ''

        self._parent: FileNode | None = parent
        self._filename_or_path: str = os.path.normpath(filename_or_path) # `normpath` also strips trailing slash if present
        self._is_ads: bool = is_ads

        self._depth: int = 0 if self.is_root else self.parent.depth + 1

        self._progress_report_children: OrderedDict[str, FileNode | None] = OrderedDict()
        self._progress_report_completed: bool = False
        self._progress_report_descendant_error: bool = False


        # Set tree-wide globals

        self._root: FileNode
        self._hard_links: dict[str, list[str]]
        self._hard_link_targets: dict[str, FileNode]

        if self.is_root:
            self._root = self
            self._hard_links = {}
            self._hard_link_targets = {}

            # Root only
            self._build_timestamp: float = time.time()
            self._last_progress_report_time: float | None = None
            self._should_report_progress: bool = should_report_progress
        else:
            self._root = self.parent._root
            self._hard_links = self.parent._hard_links
            self._hard_link_targets = self.parent._hard_link_targets


        # Add this node to the parent node's progress report display

        # noinspection PyUnresolvedReferences
        if self.root._should_report_progress and not self.is_root:
            # noinspection PyUnresolvedReferences
            self.parent._progress_report_children[self.filename_or_path] = self


        self._error: Exception | None = None
        progress_report_was_updated: bool = False

        try:

            # Get file path and canonical path

            path: str = self.path
            canonical_path: str = os.path.realpath(path, strict=True)


            # Get file type

            self._file_type: FileType

            if self.is_ads:
                self._file_type = FileType.ALTERNATE_DATA_STREAM
            elif os.path.islink(path):
                self._file_type = FileType.SYMLINK
            elif os.path.isjunction(path):
                self._file_type = FileType.JUNCTION
            elif canonical_path in self._hard_links:
                self._file_type = FileType.HARD_LINK
            elif os.path.isdir(path):
                self._file_type = FileType.DIRECTORY
            elif os.path.isfile(path):
                self._file_type = FileType.REGULAR_FILE
            else:
                self._file_type = FileType.UNKNOWN


            # Update progress report

            self._report_progress_if_needed()
            progress_report_was_updated = True


            # Get all names (hard links) of the file as absolute paths, and get its hard link target (the first
            # discovered node that is hard linked to it, and is not labeled a hard link because it is the first)

            self._hard_link_target: FileNode

            if self.file_type is FileType.HARD_LINK:
                # The node has multiple names and is not the first discovered, so it is labeled a hard link.
                self._hard_link_target = self._hard_link_targets[canonical_path]

                # Because this name cluster has already been discovered, the list of names is already stored in
                # `self._hard_links[path]`, so there is no need to set it.

            elif self.file_type in (FileType.ALTERNATE_DATA_STREAM, FileType.SYMLINK, FileType.JUNCTION):
                # If the node is an alternate data stream, hard links are impossible at this level; they operate
                # per-file, not per-stream.
                #
                # If the node is a symlink or junction, it cannot also have multiple names, i.e. it cannot also be a
                # hard link.

                self._hard_link_target = self

            else:
                # The node has one name or is the first of multiple names to be discovered, so it is not labeled a hard
                # link.
                self._hard_link_target = self

                # `win32file.FindFileNames` returns a list of absolute paths, but they do not include a drive letter. We
                # fix this by joining with the drive root of the node's canonical path.
                canonical_drive_root: str = os.path.splitdrive(canonical_path)[0] + os.path.sep
                canonical_names: list[str] = win32file.FindFileNames(canonical_path)
                all_paths: list[str] = list(map(lambda name: os.path.join(canonical_drive_root, name), canonical_names))

                # If there are multiple names, map each of them to the list of names and a reference to `self` so that
                # any single name can be used to retrieve them later.
                if len(all_paths) > 1:
                    for path_ in all_paths:
                        self._hard_links[path_] = all_paths
                        self._hard_link_targets[path_] = self


            # File metadata attributes

            self._logical_size: int
            self._physical_size: int
            self._total_logical_size: int
            self._total_physical_size: int
            self._is_drive_root: bool
            self._drive_capacity: int | None # Only set for drive roots, and `None` if we do not have admin privileges
            self._drive_free_space: int # Only set for drive roots
            self._windows_file_attributes: WindowsFileAttributes


            # Get file size

            if self.can_store_data:
                # The file is a regular file or alternate data stream, so get its logical and physical size

                # Get logical size
                self._logical_size = os.path.getsize(path)

                # Get physical size
                #
                # This took me a while to figure out, but it works!!
                # noinspection PyTypeChecker
                handle: pywintypes.HANDLEType = win32file.CreateFile(
                    path, # fileName
                    0, # desiredAccess
                    win32file.FILE_SHARE_READ, # shareMode
                    None, # attributes
                    win32file.OPEN_EXISTING, # CreationDisposition
                    0, # flagsAndAttributes
                    None, # hTemplateFile
                )
                try:
                    # noinspection PyTypeChecker
                    file_standard_info: dict[str, int | bool] = win32file.GetFileInformationByHandleEx(
                        handle, # File
                        win32file.FileStandardInfo, # FileInformationClass
                    )
                    self._physical_size = file_standard_info['AllocationSize']
                finally:
                    handle.close()

                # Another noteworthy function: `win32file.GetCompressedFileSize(path)` appears to return the size on
                # disk of the file data, but not including other parts like metadata. The only case where this differs
                # from regular file size are in so-called "compressed" and "sparse" files. The former can be created
                # with a checkbox in the "Advanced..." menu of a file's properties in File Explorer.

            else:
                # The file does not intrinsically store data

                self._logical_size = 0
                self._physical_size = 0

            # Any child nodes that are created later will add their own logical and physical file size to this node's
            # respective totals.
            self._total_logical_size = self._logical_size
            self._total_physical_size = self._physical_size


            # Get drive capacity and free space

            self._is_drive_root: bool

            if self.is_root:
                # Check if the node is a drive root
                drive_letter: str; path_on_drive: str
                drive_letter, path_on_drive = os.path.splitdrive(canonical_path)
                self._is_drive_root = drive_letter != '' and path_on_drive == os.path.sep

                # If it is, get drive size info
                if self._is_drive_root:
                    # Get drive capacity (requires administrator privileges for some reason)
                    if win32com.shell.shell.IsUserAnAdmin():
                        # This also took me a little while to figure out
                        # noinspection PyTypeChecker
                        handle: pywintypes.HANDLEType = win32file.CreateFile(
                            fr'\\.\{drive_letter}', # fileName; e.g. `\\.\C:`
                            win32file.GENERIC_READ, # desiredAccess
                            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE, # shareMode
                            None, # attributes
                            win32file.OPEN_EXISTING, # CreationDisposition
                            0, # flagsAndAttributes
                            None, # hTemplateFile
                        )
                        try:
                            # noinspection PyTypeChecker
                            drive_size_buffer: bytes = win32file.DeviceIoControl(
                                handle, # Device
                                winioctlcon.IOCTL_DISK_GET_LENGTH_INFO, # IoControlCode
                                None, # InBuffer
                                8, # OutBuffer
                            )
                            self._drive_capacity = int.from_bytes(drive_size_buffer, byteorder='little')
                        finally:
                            handle.close()
                    else:
                        # `None` indicates insufficient permissions
                        self._drive_capacity = None

                    # Get drive free space
                    user_available_free_space: int; user_available_capacity: int; free_space: int
                    user_available_free_space, user_available_capacity, free_space = win32api.GetDiskFreeSpaceEx(canonical_path)
                    self._drive_free_space = free_space

            else:
                # Non-root nodes are assumed not to be a drive root
                self._is_drive_root = False


            # Get Windows file attributes

            self._windows_file_attributes = WindowsFileAttributes(win32file.GetFileAttributes(path))


            # We may only access attributes of the parent node in this function that are defined above this point.


            # Create child nodes

            self._children: dict[str, FileNode]

            if self.file_type is FileType.DIRECTORY:
                # Create a child node for each file in the directory

                # List directory contents
                child_filenames: list[str] = os.listdir(path)

                # Add each file to progress report display; the child nodes will add themselves in place of `None` once
                # created
                # noinspection PyUnresolvedReferences
                if self.root._should_report_progress:
                    for filename in child_filenames:
                        self._progress_report_children[filename] = None

                # Create child nodes
                self._children = {filename: FileNode(
                    filename,
                    parent=self,
                    catch_errors=catch_errors,
                ) for filename in child_filenames}

            elif self.file_type is FileType.REGULAR_FILE:
                # Create a child node for each alternate data stream

                self._children = {}
                filename: str = self.filename

                # Find file streams
                streams: list[tuple[int, str]] = win32file.FindStreams(path)

                # Filter for alternate data streams
                alternate_data_streams: list[str] = []
                for stream_file_size, stream_suffix in streams:
                    # `stream_suffix` is of the form ":<stream_name>:<stream_type>".
                    stream_name, stream_type = stream_suffix[1:].split(':')

                    # "$DATA" indicates a data stream.
                    if stream_type != '$DATA':
                        continue

                    # An empty stream name indicates the main stream, which is not an alternate stream; it is the stream
                    # that holds the main file data, which is already accounted for.
                    if stream_name == '':
                        continue

                    alternate_data_streams.append(stream_suffix)

                # Add each alternate data stream name to progress report display; the child nodes will add themselves in
                # place of `None` once created
                # noinspection PyUnresolvedReferences
                if self.root._should_report_progress:
                    for stream_suffix in alternate_data_streams:
                        self._progress_report_children[filename + stream_suffix] = None

                # Create child nodes
                for stream_suffix in alternate_data_streams:
                    self._children[stream_suffix] = FileNode(
                        filename + stream_suffix,
                        parent=self,
                        is_ads=True,
                        catch_errors=catch_errors,
                    )

            else:
                # Other types do not store children
                self._children = {}

            # No need to display children in progress report anymore once they are all initialized
            # noinspection PyUnresolvedReferences
            if self.root._should_report_progress:
                self._progress_report_children = OrderedDict()


        # Error handling

        except (OSError, pywintypes.error) as error:
            self._error = error

            # Re-raise error if catching errors is disabled
            if not catch_errors:
                raise

            # Set all attributes defined in the `try` block to their most empty state
            self._hard_link_target = self
            self._file_type = FileType.UNKNOWN
            self._logical_size = 0
            self._physical_size = 0
            self._total_logical_size = 0
            self._total_physical_size = 0
            self._is_drive_root = False
            self._windows_file_attributes = WindowsFileAttributes(0)
            self._progress_report_children = OrderedDict()
            self._children = {}

            # Update progress report if the error occurred before it could be updated normally
            if not progress_report_was_updated:
                self._report_progress_if_needed()

        finally:
            # We no longer need a reference to the hard link names map
            del self._hard_links


        # Merge total file size into parent
        #
        # This is only done after creating the child nodes: they will in turn merge their file sizes into to this node
        # at that point, and we only want to merge our total once it is final.
        if not self.is_root:
            # noinspection PyUnresolvedReferences
            self.parent._total_logical_size += self.total_logical_size
            # noinspection PyUnresolvedReferences
            self.parent._total_physical_size += self.total_physical_size


        # Propagate error status to parent's progress report
        #
        # This is done after creating the child nodes for the same reason as above.
        # noinspection PyUnresolvedReferences
        if self.root._should_report_progress and not self.is_root:
            if not self.parent._progress_report_descendant_error:
                self.parent._progress_report_descendant_error = self._progress_report_descendant_error or (self._error is not None)


        # Mark progress report as completed
        # noinspection PyUnresolvedReferences
        if self.root._should_report_progress:
            self._progress_report_completed = True

            # Clear progress report display if this is the root node because once the root node finishes initializing,
            # the entire tree must've finished initializing.
            if self.is_root:
                FileNode._clear_progress_report_display()

    def __repr__(self) -> str:
        depth_info: str = 'Root' if self.is_root else f'Depth {self.depth}'

        if self.is_error:
            return f'<{type(self).__name__}: {depth_info} {self.filename_or_path!r}, ERROR>'

        size_info: str = f', log. {format_data_size(self.logical_size)}, phys. {format_data_size(self.physical_size)}' if self.can_store_data else ''
        return f'<{type(self).__name__}: {depth_info} {self.file_type.short_human_readable_name} {self.filename_or_path!r}{size_info}>'

    @property
    def filename_or_path(self) -> str:
        """
        The filename of the node, or its absolute path if it is a root node.

        See also: ``FileNode.filename``, ``FileNode.path``

        :rtype: str
        """
        return self._filename_or_path

    @property
    def filename(self) -> str:
        """
        The filename of the node (excluding parent directories even if it is a root node), computed on demand.

        See also: ``FileNode.filename_or_path``

        :rtype: str
        """
        return os.path.basename(self._filename_or_path) if self.is_root else self._filename_or_path

    @property
    def path(self) -> str:
        """
        The absolute path of the node, computed on demand.

        See also: ``FileNode.filename_or_path``

        :rtype: str
        """
        # If this is a root node, the absolute path is already known.
        if self.is_root:
            return self._filename_or_path

        # If this is an alternate data stream, replace the filename of the parent node in its absolute path with this
        # node's filename.
        if self.is_ads:
            return os.path.join(os.path.dirname(self._parent.path), self._filename_or_path)

        # Otherwise, join the filename with the absolute path of the parent node.
        return os.path.join(self._parent.path, self._filename_or_path)

    @property
    def error(self) -> Exception | None:
        """
        The error that occurred if the node failed to be initialized, otherwise ``None``.

        :rtype: Exception | None
        """
        return self._error

    @property
    def is_error(self) -> bool:
        """
        ``True`` if the node failed to be initialized.

        See also: ``FileNode.error``

        :rtype: bool
        """
        return self._error is not None

    @property
    def file_type(self) -> FileType:
        """
        The file type of the node.

        :rtype: FileType
        """
        return self._file_type

    @property
    def is_ads(self) -> bool:
        """
        ``True`` if the node is an alternate data stream.

        :rtype: bool
        """
        return self._is_ads

    @property
    def can_store_data(self) -> bool:
        """
        ``True`` if the node can intrinsically store data.

        This is only true for regular files and alternate data streams.

        See also: ``FileType.can_store_data``

        :rtype: bool
        """
        return self._file_type.can_store_data

    @property
    def parent(self) -> FileNode:
        """
        The node's parent node.

        See also: ``FileNode.ancestors_iter``, ``FileNode.depth``

        :rtype: FileNode
        :raises ValueError: If the node is a root node, i.e. it has no parent.
        """
        if self._parent is None:
            raise ValueError('Root node has no parent')
        return self._parent

    @property
    def children(self) -> dict[str, FileNode]:
        """
        The node's child nodes in a dictionary.

        - For ``FileType.DIRECTORY`` nodes, the dictionary is keyed by filename.
        - For ``FileType.REGULAR_FILE`` nodes, the dictionary is keyed by alternate data stream suffix of the form ``:<stream_name>:$DATA``.
        - For other nodes, the dictionary is empty.

        Do not mutate this dictionary externally.

        See also: ``FileNode.descendants_iter``

        :rtype: dict[str, FileNode]
        """
        return self._children

    @property
    def depth(self) -> int:
        """
        The depth of the node, i.e. its number of ancestors.

        See also: ``FileNode.ancestors_iter``

        :return: The depth of the node.
        :rtype: int
        """
        return self._depth

    @property
    def root(self) -> FileNode:
        """
        The root node of the file tree.

        :rtype: FileNode
        """
        return self._root

    @property
    def hard_link_target(self) -> FileNode:
        """
        The first node discovered that is hard linked to ``self``, and therefore does not have the type
        ``FileType.HARD_LINK``.

        If ``self.file_type`` is not ``FileType.HARD_LINK``, returns ``self``.

        :rtype: FileNode
        """
        return self._hard_link_target

    @property
    def is_root(self) -> bool:
        """
        ``True`` if the node is a root node, i.e. it has no parent.

        :rtype: bool
        """
        return self._parent is None

    @property
    def is_leaf(self) -> bool:
        """
        ``True`` if the node is a leaf node, i.e. it has no children.

        :rtype: bool
        """
        return len(self._children) == 0

    @property
    def logical_size(self) -> int:
        """
        The logical size of the node in bytes (not including descendants).

        This is the size of the represented data as opposed to the size on disk. For the latter, use
        ``FileNode.physical_size``.

        Nodes with a file type other than ``FileType.REGULAR_FILE`` or ``FileType.ALTERNATE_DATA_STREAM`` are considered
        to have a logical size of zero to avoid double counting.

        If you wish to include descendants, use ``FileNode.total_logical_size`` instead.

        :rtype: int
        """
        return self._logical_size

    @property
    def physical_size(self) -> int:
        """
        The physical size of the node in bytes, i.e. its size on disk (not including descendants).

        For the size of the represented data, use ``FileNode.logical_size`` instead.

        If you want to include descendants, use ``FileNode.total_physical_size`` instead.

        :rtype: int
        """
        return self._physical_size

    @property
    def total_logical_size(self) -> int:
        """
        The total logical size of the node and its descendants in bytes.

        This is the size of the represented data as opposed to the size on disk. For the latter, use
        ``FileNode.total_physical_size``.

        Nodes with a file type other than ``FileType.REGULAR_FILE`` or ``FileType.ALTERNATE_DATA_STREAM`` are considered
        to have a logical size of zero to avoid double counting.

        If you do not wish to include descendants, use ``FileNode.logical_size`` instead.

        :rtype: int
        """
        return self._total_logical_size

    @property
    def total_physical_size(self) -> int:
        """
        The total physical size of the node and its descendants in bytes, i.e. their size on disk.

        For the size of the represented data, use ``FileNode.total_logical_size`` instead.

        If you do not wish to include descendants, use ``FileNode.physical_size`` instead.

        :rtype: int
        """
        return self._total_physical_size

    @property
    def is_drive_root(self) -> bool:
        """
        ``True`` if the node is a drive root.

        :rtype: bool
        """
        return self._is_drive_root

    @property
    def drive_capacity(self) -> int | None:
        """
        The capacity of the drive in bytes (if the node is a drive root), or ``None`` if the node was not initialized
        with administrator privileges.

        :rtype: int | None
        """
        assert self._is_drive_root
        return self._drive_capacity

    @property
    def drive_free_space(self) -> int:
        """
        The free space on the drive in bytes (if the node is a drive root).

        :rtype: int
        """
        assert self._is_drive_root
        return self._drive_free_space

    @property
    def drive_used_space(self) -> int:
        """
        The used space on the drive in bytes (if the node is a drive root), or ``None`` if the node was not initialized
        with administrator privileges.

        :rtype: int | None
        """
        assert self._is_drive_root
        return None if self._drive_capacity is None else self._drive_capacity - self._drive_free_space

    @property
    def windows_file_attributes(self) -> WindowsFileAttributes:
        """
        The Windows file attributes of the node.

        :rtype: WindowsFileAttributes
        """
        return self._windows_file_attributes

    @property
    def color_rgb(self) -> tuple[int, int, int]:
        """
        A color in RGB format to display the node.

        Mostly equivalent to ``self.file_type.color_rgb``, except there is a separate color if an error occurred (i.e.
        if ``self.is_error`` is true).

        :rtype: tuple[int, int, int]
        """
        return ERROR_COLOR if self.is_error else self.file_type.color_rgb

    @property
    def build_timestamp(self) -> float:
        """
        The Unix timestamp of when the file tree was built.

        :rtype: float
        """
        return self.root._build_timestamp

    def _format_progress(self, working_node: FileNode) -> list[str]:
        # All lines are assumed to begin with default formatting.

        lines: list[str] = []
        r: int; g: int; b: int

        # Add loading text if this is the root node
        if self.is_root:
            lines.append('Building tree...')


        # Format header line

        header_line: str = ''

        # Add checkbox
        if self._progress_report_completed:
            if self.is_error:
                r, g, b = ERROR_COLOR
                header_line += f'[\033[38;2;{r};{g};{b}mX\033[0m] '
            elif self._progress_report_descendant_error:
                r, g, b = DESCENDANT_ERROR_COLOR
                header_line += f'[\033[38;2;{r};{g};{b}m/\033[0m] '
            else:
                r, g, b = COMPLETED_COLOR
                header_line += f'[\033[38;2;{r};{g};{b}m\u2713\033[0m] '
        elif self is working_node:
            header_line += '[\u2026] '
        else:
            header_line += '[ ] '

        # Color according to file type or error
        r, g, b = self.color_rgb
        header_line += f'\033[38;2;{r};{g};{b}m'

        # Add filename or path
        header_line += truncate(self.filename_or_path, PROGRESS_REPORT_MAX_FILENAME_WIDTH)
        lines.append(header_line)


        # Truncate list of children

        # Get the ancestor of the working node with the same depth as this node's children (if there is one)
        working_node_ancestor: FileNode | None = None
        if working_node.depth > self.depth:
            for ancestor_node in working_node.ancestors_iter(include_self=True):
                if ancestor_node.depth == self.depth + 1:
                    working_node_ancestor = ancestor_node
                    break

        # Get index of child to keep on screen (center of visible children)
        keep_child_index_on_screen: int = 0
        if working_node_ancestor is not None:
            for i, child_node in enumerate(self._progress_report_children.values()):
                if child_node is working_node_ancestor:
                    keep_child_index_on_screen = i
                    break

        # Get minimum and maximum visible child indices
        max_visible_children: int = PROGRESS_REPORT_WORKING_DEPTH_MAX_VISIBLE_CHILDREN if self.depth + 1 == working_node.depth else PROGRESS_REPORT_OUTER_MAX_VISIBLE_CHILDREN
        min_visible_child_index: int = max(0, min(keep_child_index_on_screen - math.floor((max_visible_children - 1) / 2), len(self._progress_report_children) - max_visible_children))
        max_visible_child_index: int = min(max(max_visible_children - 1, keep_child_index_on_screen + math.ceil((max_visible_children - 1) / 2)), len(self._progress_report_children) - 1)

        # Truncate list of children
        truncated_children: list[tuple[str, FileNode | None]] = []
        for i, (filename, child_node) in enumerate(self._progress_report_children.items()):
            if i < min_visible_child_index: continue
            if i > max_visible_child_index: break
            truncated_children.append((filename, child_node))

        # Check whether each end of the list was actually truncated
        child_list_start_was_truncated: bool = min_visible_child_index > 0
        child_list_end_was_truncated: bool = max_visible_child_index < len(self._progress_report_children) - 1


        # Format child node lines

        # Add dotted line if start of list was truncated
        if child_list_start_was_truncated:
            lines.append(' \u250a')

        # Format each child node
        for i, (filename, child_node) in enumerate(truncated_children):
            is_last_child: bool = (i + min_visible_child_index) == len(self._progress_report_children) - 1
            branching_symbol: str = '\u2570' if is_last_child else '\u251c'

            if child_node is None:
                # Node has not been created yet

                line: str = ''

                # Add branching symbol
                line += ' ' + branching_symbol

                # Add checkbox
                line += '[ ] '

                # Color
                r, g, b = NOT_STARTED_COLOR
                line += f'\033[38;2;{r};{g};{b}m'

                # Add filename
                line += truncate(filename, PROGRESS_REPORT_MAX_FILENAME_WIDTH)

                lines.append(line)

            else:
                # Node has been created

                # Format its progress report
                child_lines: list[str] = child_node._format_progress(working_node=working_node)

                # Add branching symbol
                child_lines[0] = ' ' + branching_symbol + child_lines[0]

                # Add branching line symbols to all non-header lines
                for j in range(1, len(child_lines)):
                    child_lines[j] = ('  ' if is_last_child else ' \u2502') + child_lines[j]

                lines.extend(child_lines)

        # Add dotted line if end of list was truncated
        if child_list_end_was_truncated:
            lines.append(' \u250a')


        return lines

    def _report_progress(self, working_node: FileNode) -> None:
        lines: list[str] = self._format_progress(working_node=working_node)
        progress_report: str = '\n'.join('\033[0m' + line for line in lines) # Each line must start with default formatting

        output: str = ''
        output += '\033[0J' # Clear from cursor until end of screen
        output += progress_report # Print progress report
        output += '\033[G' # Move cursor to leftmost column
        output += '\033[F' * (len(lines) - 1) # Move cursor up to start of first line
        sys.stdout.write(output)
        sys.stdout.flush()

    @staticmethod
    def _clear_progress_report_display() -> None:
        output: str = ''
        output += '\033[0J' # Clear from cursor until end of screen
        output += '\033[0m' # Reset formatting
        sys.stdout.write(output)
        sys.stdout.flush()

    def _report_progress_if_needed(self) -> None:
        # Check if progress should be reported
        if not self.root._should_report_progress:
            return
        now: float = time.time()
        last_report_time: float | None = self.root._last_progress_report_time
        if last_report_time is not None and now - last_report_time < PROGRESS_REPORT_INTERVAL:
            return
        self.root._last_progress_report_time = now

        # Report progress
        self.root._report_progress(working_node=self)

    def ancestors_iter(self, include_self: bool = False) -> Iterator[FileNode]:
        """
        Finds all the ancestors of the node.

        See also: ``FileNode.descendants_iter``

        :param include_self: If ``True``, the node itself will be included in the result.
        :type include_self: bool
        :return: An iterator that yields the node's ancestors in order of decreasing depth.
        :rtype: Iterator[FileNode]
        """
        node: FileNode = self

        # Yield self
        if include_self:
            yield self

        # Keep moving up one level until we reach the root node
        while not node.is_root:
            node = node.parent
            yield node

    def descendants_iter(
            self,
            max_depth: int | None = None,
            data_only: bool = False,
            files_only: bool = False,
            leaf_only: bool = False,
            include_self: bool = False,
    ) -> Iterator[FileNode]:
        """
        Finds all the descendants of the node.

        See also: ``FileNode.ancestors_iter``, ``FileNode.can_store_data``

        :param max_depth: The maximum depth of the search; only descendants that are less than or equal to this depth
            (relative to the starting node) will be included. If ``None``, the maximum depth is unlimited.
        :type max_depth: int | None
        :param data_only: If ``True``, the result is filtered to only ``FileType.REGULAR_FILE`` and
            ``FileType.ALTERNATE_DATA_STREAM`` nodes, i.e. nodes that can intrinsically store data.
        :type data_only: bool
        :param files_only: If ``True``, the result is filtered to only ``FileType.REGULAR_FILE`` nodes.
        :type files_only: bool
        :param leaf_only: If ``True``, the result is filtered to only leaf nodes (nodes that have no children).
        :type leaf_only: bool
        :param include_self: If ``True``, the node itself will be included in the result (if the other filters allow
            it).
        :type include_self: bool
        :return: An iterator that yields the node's descendants in an arbitrary order.
        :rtype: Iterator[FileNode]
        """
        # Quit if maximum depth exceeded
        if max_depth is not None and max_depth < 0:
            return

        # Yield self
        if include_self \
            and ((not data_only) or self.can_store_data) \
            and ((not files_only) or self.file_type is FileType.REGULAR_FILE) \
            and ((not leaf_only) or self.is_leaf):
            yield self

        # No need to recurse to children if we are at the maximum depth.
        #
        # This guard clause is not logically necessary, as there is already one at the start of the function that
        # catches strictly more* than this one does; this guard clause is just for efficiency, to prevent unnecessary
        # looping and calls.
        #
        # *It also catches cases where the `max_depth` parameter of the outermost call is negative, whereas this clause
        # does not, because this clause would not catch those cases until after `self` has already been yielded.
        if max_depth is not None and max_depth <= 0:
            return

        # Recursively yield descendants of children
        for identifier, child_node in self.children.items():
            yield from child_node.descendants_iter(
                max_depth = None if max_depth is None else max_depth - 1,
                data_only = data_only,
                files_only = files_only,
                leaf_only = leaf_only,
                include_self = True,
            )

    def is_ancestor(self, other_node: FileNode, include_self: bool = True) -> bool:
        """
        Checks whether this node is an ancestor of another node.

        :param other_node: The node to check that this is an ancestor of.
        :type other_node: FileNode
        :param include_self: If ``True``, the node itself is also considered an ancestor.
        :type include_self: bool
        :return: ``True`` if ``self`` is an ancestor of ``other_node``.
        :rtype: bool
        """
        # Move up the file tree starting from the other node until we reach or pass the depth of this node
        node: FileNode = other_node
        while self.depth < node.depth:
            node = node.parent

        # This node is an ancestor if it is the node we just found
        return node is self and (include_self or node is not other_node)

    @staticmethod
    def deepest_common_ancestor(*nodes: FileNode) -> FileNode:
        """
        Finds the deepest common ancestor of a group of nodes.

        :param nodes: The nodes to find the deepest common ancestor of. Must be non-empty.
        :type nodes: FileNode
        :return: The deepest common ancestor of all the nodes given.
        :rtype: FileNode
        :raises ValueError: If ``nodes`` is empty, or if not all the nodes are on the same tree. The latter can be
            checked using ``FileNode.has_common_root`` if desired.
        """
        if len(nodes) == 0:
            raise ValueError('\'nodes\' must be non-empty')

        # Find the depth of the outermost node; the common ancestor must be at most this deep
        max_depth: int = min(map(lambda node: node.depth, nodes))

        # Find the ancestor of each node at that depth
        ancestors: list[FileNode] = []
        for node in nodes:
            for ancestor in node.ancestors_iter(include_self=True):
                if ancestor.depth == max_depth:
                    ancestors.append(ancestor)
                    break

        # Move up the file tree until all the ancestors are the same
        while True:
            # Check if all ancestors are the same
            common_ancestor: FileNode | None = ancestors[0]
            for ancestor in ancestors[1:]:
                if ancestor is not common_ancestor:
                    common_ancestor = None
                    break

            # Finish if they are
            if common_ancestor is not None:
                return common_ancestor

            # Otherwise ensure none of them are root
            if any(map(lambda node: node.is_root, ancestors)):
                raise ValueError('All nodes must be on the same file tree')

            # And then move up the file tree
            ancestors = [ancestor.parent for ancestor in ancestors]

    @staticmethod
    def has_common_root(*nodes: FileNode) -> bool:
        """
        Checks if a group of nodes are on the same tree, i.e. they have a common root.

        :param nodes: The nodes to check if they are on the same tree.
        :type nodes: FileNode
        :return: ``True`` if all the nodes are on the same tree (including if ``nodes`` is empty).
        :rtype: bool
        """
        # The nodes are considered to be on the same tree if no nodes are given.
        if len(nodes) == 0:
            return True

        # Check if all nodes have the same root as the first node
        common_root: FileNode = nodes[0].root
        for node in nodes[1:]:
            if node.root is not common_root:
                return False

        return True

    def create_flame_graph(
            self,
            use_physical_size: bool = True,
            max_depth: int | None = None,
            width: int = 1920,
            layer_height: int = 20,
            background_color: tuple[int, int, int, int] = (255, 255, 255, 255),
            foreground_color: tuple[int, int, int, int] = (0, 0, 0, 255),
            label_visibility: LabelVisibility = LabelVisibility.ALL,
            min_label_width: int = 15,
            hide_full_root_path: bool = False,
            font: ImageFont.ImageFont | ImageFont.FreeTypeFont | None = None,
            show_drive_free_space: bool = True,
            show_drive_unaccounted_space: bool = True,
            show_drive_extra_counted_space: bool = True,
    ) -> Image.Image:
        """
        Creates a flame graph visualization of the file tree starting from this node.

        Each "layer" of the graph corresponds to a specific node depth, with the "root layer" being this node's depth.

        :param use_physical_size: If ``True``, the physical size of nodes will be displayed. If ``False``, their logical
            size will be displayed.
        :type use_physical_size: bool
        :param max_depth: Nodes deeper than this depth (relative to the root layer) will not be drawn. If ``None``,
            there is no depth limit.
        :type max_depth: int | None
        :param width: The width of the graph in pixels. Must be greater than ``1``.
        :type width: int
        :param layer_height: The height of each layer in pixels. Must be greater than ``1``.
        :type layer_height: int
        :param background_color: The background color of the graph in RGBA format.
        :type background_color: tuple[int, int, int, int]
        :param foreground_color: The color of the graph's text and outlines in RGBA format.
        :type foreground_color: tuple[int, int, int, int]
        :param label_visibility: Determines where to draw labels on rectangles. ``LabelVisibility.ALL`` will always draw
            labels. ``LabelVisibility.SPECIAL_SEGMENTS`` will only draw labels on the rectangles created by
            ``show_drive_free_space``, ``show_drive_unaccounted_space``, and ``show_drive_extra_counted_space``.
            ``LabelVisibility.FILES`` will draw labels on all rectangles except those that have labels with
            ``LabelVisibility.SPECIAL_SEGMENTS``, i.e. all nodes of the file tree. ``LabelVisibility.NONE`` will never
            draw labels.
        :type label_visibility: bool
        :param min_label_width: If a rectangle is less than this many pixels wide, its label will not be drawn. Note
            that lower values may take longer to render due to increased overall label count. Must be at least ``1``.
        :type min_label_width: int
        :param hide_full_root_path: If ``True``, only the filename part of the node's path is shown on its label. Has no
            effect on drive roots.
        :type hide_full_root_path: bool
        :param font: The font to use for node labels. If ``None``, a default font is used.
        :type font: ImageFont.ImageFont | ImageFont.FreeTypeFont | None
        :param show_drive_free_space: Only applies if the node is a drive root and ``use_physical_size`` is true. If
            ``True``, a rectangle will be drawn at the root layer representing the amount of free space on the drive.
            Extra-counted space is not included (if the data to compute that is accessible); see
            ``show_drive_extra_counted_space`` for info about extra-counted space.
        :type show_drive_free_space: bool
        :param show_drive_unaccounted_space: Only applies if the node is a drive root and ``use_physical_size`` is true.
            If ``True``, a rectangle will be drawn at the root layer representing the amount of extra used space on the
            drive that was not accounted for by the other nodes on the graph. For example, if Windows reports the drive
            to have a 400 GB capacity with 100 GB of free space, but the combined size of the nodes on the graph is only
            250 GB, the unaccounted space would be 50 GB. In practice, this should include files that the program failed
            to retrieve the metadata of. If the file tree object structure was not created with administrator
            privileges, the necessary data to compute this is inaccessible, and so this will not be drawn.
        :type show_drive_unaccounted_space: bool
        :param show_drive_extra_counted_space: Only applies if the node is a drive root and ``use_physical_size`` is
            true. If ``True``, a rectangle will be drawn at the root layer representing the amount of space used by
            nodes on the graph that shouldn't have been used because it overlaps the free space on the drive. For
            example, if Windows reports the drive to have a 400 GB capacity with 100 GB of free space, but the combined
            size of the nodes on the graph is 350 GB, the extra-counted space would be 50 GB. If the file tree object
            structure was not created with administrator privileges, the necessary data to compute this is inaccessible,
            and so this will not be drawn.
        :type show_drive_extra_counted_space: bool
        :return: The flame graph image.
        :rtype: Image.Image
        """
        # Compute drive free, unaccounted, and extra-counted space
        #
        # `None` means the respective statistic will not be displayed.

        drive_free_space: int | None = None
        drive_unaccounted_space: int | None = None # Negative values indicate extra-counted space

        if self.is_drive_root and use_physical_size:
            # Free space
            if show_drive_free_space:
                drive_free_space = self.drive_free_space

                # Cap to not include extra-counted space
                if self.drive_capacity is not None:
                    drive_free_space = max(0, min(drive_free_space, self.drive_capacity - self.total_physical_size))

            # Unaccounted/extra-counted space
            if self.drive_capacity is not None:
                # Compute
                drive_unaccounted_space = self.drive_used_space - self.total_physical_size

                # Check if the statistic should be shown
                is_shown: bool = show_drive_unaccounted_space if drive_unaccounted_space >= 0 else show_drive_extra_counted_space
                if not is_shown:
                    drive_unaccounted_space = None


        # Compute total graph size in bytes
        graph_size_bytes: int = self.total_physical_size if use_physical_size else self.total_logical_size
        if drive_free_space is not None:
            graph_size_bytes += drive_free_space
        if drive_unaccounted_space is not None:
            graph_size_bytes += abs(drive_unaccounted_space)

        # Get base depth (depth of root layer)
        base_depth: int = self.depth

        # Get number of layers
        layer_count: int = 1
        for node in self.descendants_iter(include_self=True):
            # Compute each node's depth relative to the root layer
            relative_depth: int = node.depth - base_depth

            if relative_depth >= layer_count:
                # Update the layer count to the depth of the deepest node plus one (to include root layer)
                layer_count = relative_depth + 1

                # Enforce maximum depth if set
                #
                # If the maximum depth is reached, we can also exit the loop because there is no way we could find a
                # larger acceptable depth.
                if max_depth is not None and relative_depth >= max_depth:
                    layer_count = max_depth + 1
                    break

        # Compute image dimensions
        height: int = layer_height * layer_count + 1 # Extra pixel for outline of rectangles on the top layer
        pixels_per_byte: float = (width - 1) / graph_size_bytes # 1 pixel subtracted here for outline of rectangles on the right edge

        # Create image
        im: Image.Image = Image.new('RGBA', (width, height), background_color)
        draw: ImageDraw.ImageDraw = ImageDraw.Draw(im)

        # Load default font if a font is not provided
        if font is None:
            font = ImageFont.truetype('C:\\Windows\\Fonts\\bahnschrift.ttf', size=min(10.0, layer_height * 0.6))
        
        # Define function to recursively draw nodes
        def draw_node(node: FileNode | tuple[str, int, tuple[int, int, int, int]], horizontal_bytes_offset: int) -> int:
            """
            Draws a node and its descendants or a special segment on the flame graph.

            "Special segments" are drawn like a node at the root layer and have no children.

            :param node: If drawing a node, the node to draw. Otherwise, a ``(label, size, color)`` tuple where
                ``label`` is the text of the label, ``size`` is the width of the special segment in bytes, and ``color``
                is the color of the special segment in RGBA format.
            :type node: FileNode | tuple[str, int, tuple[int, int, int, int]]
            :param horizontal_bytes_offset: The number of bytes to offset the node/special segment horizontally in the
                graph.
            :type horizontal_bytes_offset: int
            :return: The width of the node/special segment in bytes.
            :rtype: int
            """
            nonlocal use_physical_size
            nonlocal max_depth
            nonlocal layer_height
            nonlocal foreground_color
            nonlocal label_visibility
            nonlocal min_label_width
            nonlocal font

            nonlocal base_depth
            nonlocal layer_count
            nonlocal pixels_per_byte
            nonlocal im
            nonlocal draw

            nonlocal draw_node

            # Get node/special segment properties
            is_node: bool = isinstance(node, FileNode)
            node_depth: int # Relative to the root layer
            node_size_bytes: int
            label_text: str
            if is_node:
                node_depth = node.depth - base_depth
                node_size_bytes = node.total_physical_size if use_physical_size else node.total_logical_size
                label_text = (
                    node.path
                        if (node_depth == 0 and not hide_full_root_path) or node.is_drive_root
                            else node.filename
                )
            else:
                node_depth = 0
                node_size_bytes = node[1]
                label_text = node[0]
            
            # Do not draw if the node is above the maximum depth
            if max_depth is not None and node_depth > max_depth:
                return node_size_bytes

            # Compute rectangle (north-west corner) position
            rectangle_x: float = horizontal_bytes_offset * pixels_per_byte
            rectangle_y: int = (layer_count - node_depth - 1) * layer_height

            # Compute rectangle width
            rectangle_width: float = node_size_bytes * pixels_per_byte

            # Get rectangle color
            rectangle_color: tuple[int, int, int, int] = node.color_rgb + (255,) if is_node else node[2]

            # Draw rectangle
            draw.rectangle(
                [
                    (round(rectangle_x), round(rectangle_y)),
                    (round(rectangle_x + rectangle_width), round(rectangle_y + layer_height))
                ],
                fill=rectangle_color,
                outline=foreground_color,
                width=1,
            )

            # Draw label if labels are enabled for this rectangle type and the rectangle is big enough
            if label_visibility.is_visible(is_node) and rectangle_width >= min_label_width:

                # Create clipping canvas for label
                #
                # The clipping canvas' width is the width of the rectangle, not including half of the outline, rounded
                # down. Its height is the height of the rectangle, not including the outline.
                label_canvas: Image.Image = Image.new('RGBA', (math.floor(rectangle_width), layer_height - 1), (0, 0, 0, 0))
                label_canvas_draw: ImageDraw.ImageDraw = ImageDraw.Draw(label_canvas)

                # Compute label (center) position relative to clipping canvas
                #
                # We subtract 1 from the clipping canvas' dimensions and divide by 2 to get its center.
                label_x_on_canvas: float = (label_canvas.width - 1) / 2
                label_y_on_canvas: float = (label_canvas.height - 1) / 2

                # Draw label onto clipping canvas
                label_canvas_draw.text(
                    xy = (label_x_on_canvas, label_y_on_canvas),
                    text = label_text,
                    fill = foreground_color,
                    font = font,
                    anchor = 'mm',
                )

                # Compute clipping canvas (north-west corner) position
                label_canvas_x: int = math.ceil(rectangle_x) # Approximately on or to the right of the left outline
                label_canvas_y: int = rectangle_y + 1 # Exactly below the top outline

                # Paste clipping canvas onto graph
                im.alpha_composite(label_canvas, (label_canvas_x, label_canvas_y))

            # Special segments do not have children
            if not is_node:
                return node_size_bytes

            # No need to draw child nodes if this node is at the maximum depth.
            #
            # This statement isn't strictly necessary because there's a statement at the start of the function that
            # would catch this for each child, but checking here as well saves from performing an unnecessary loop
            # below.
            if max_depth is not None and node_depth == max_depth:
                return node_size_bytes

            # Draw child nodes
            child_horizontal_bytes_offset: int = horizontal_bytes_offset
            for child in node.children.values():
                child_horizontal_bytes_offset += draw_node(child, child_horizontal_bytes_offset)

            return node_size_bytes

        # Recursively draw nodes
        horizontal_bytes_offset: int = 0
        horizontal_bytes_offset += draw_node(self, horizontal_bytes_offset)

        # Draw special segments
        if drive_unaccounted_space is not None and drive_unaccounted_space >= 0:
            horizontal_bytes_offset += draw_node(('Unaccounted', drive_unaccounted_space, UNACCOUNTED_COLOR), horizontal_bytes_offset)
        if drive_free_space is not None:
            horizontal_bytes_offset += draw_node(('Free', drive_free_space, FREE_COLOR), horizontal_bytes_offset)
        if drive_unaccounted_space is not None and drive_unaccounted_space < 0:
            horizontal_bytes_offset += draw_node(('Extra-counted', -drive_unaccounted_space, EXTRA_COUNTED_COLOR), horizontal_bytes_offset)

        # Return image
        return im



# CONSTANTS

# Progress report general settings
PROGRESS_REPORT_INTERVAL: float = 0.1 # In seconds
PROGRESS_REPORT_MAX_FILENAME_WIDTH: int = 60 # In characters
PROGRESS_REPORT_OUTER_MAX_VISIBLE_CHILDREN: int = 5
PROGRESS_REPORT_WORKING_DEPTH_MAX_VISIBLE_CHILDREN: int = 10

# File type colors (used in flame graphs and progress reports)
FILE_TYPE_COLORS: dict[FileType, tuple[int, int, int]] = {
    FileType.DIRECTORY: (255, 150, 25),
    FileType.REGULAR_FILE: (50, 255, 50),
    FileType.ALTERNATE_DATA_STREAM: (255, 100, 255),
    FileType.SYMLINK: (0, 255, 255),
    FileType.JUNCTION: (255, 255, 0),
    FileType.HARD_LINK: (75, 100, 255),
    FileType.UNKNOWN: (200, 200, 200),
}

# Progress report colors
NOT_STARTED_COLOR: tuple[int, int, int] = (100, 100, 100)
COMPLETED_COLOR: tuple[int, int, int] = (50, 255, 50) # Only used by the checkbox; not the node itself
ERROR_COLOR: tuple[int, int, int] = (255, 50, 50)
DESCENDANT_ERROR_COLOR: tuple[int, int, int] = (255, 255, 0) # Only used by the checkbox; not the node itself

# Special segment colors
UNACCOUNTED_COLOR: tuple[int, int, int, int] = (255, 50, 50, 255)
FREE_COLOR: tuple[int, int, int, int] = (200, 200, 200, 255)
EXTRA_COUNTED_COLOR: tuple[int, int, int, int] = (255, 255, 0, 255)
