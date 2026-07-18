# This program currently only supports Windows.

# TODO: Get total remaining space on the volume using a dedicated Windows API if the root of the file tree is the root of a volume (or at least just figure out that API for later)
# TODO: Maybe delete `FileNote._hard_links` at end of `__init__`?



# IMPORTS

from __future__ import annotations

from collections.abc import Iterator
import win32file, pywintypes # pip install pywin32
import enum
import os



# CONFIG
# All of these values are currently unused.

DIRECTORY_COLOR: tuple[int, int, int] = (255, 150, 25)
REGULAR_FILE_COLOR: tuple[int, int, int] = (50, 255, 50)
ALTERNATE_DATA_STREAM_COLOR: tuple[int, int, int] = (255, 100, 255)
SYMLINK_COLOR: tuple[int, int, int] = (0, 255, 255)
JUNCTION_COLOR: tuple[int, int, int] = (255, 255, 0)
HARD_LINK_COLOR: tuple[int, int, int] = (75, 100, 255)
UNKNOWN_FILE_TYPE_COLOR: tuple[int, int, int] = (200, 200, 200)

UNACCOUNTED_COLOR: tuple[int, int, int] = (200, 200, 200)
FREE_COLOR: tuple[int, int, int] = (200, 200, 200)
ERROR_COLOR: tuple[int, int, int] = (255, 50, 50)



# DEFINITIONS

def format_data_size(size_in_bytes: int, use_iec_units: bool = True) -> str:
    """
    Formats a data size in bytes into a short human-readable string.

    :param size_in_bytes: The data size in bytes.
    :type size_in_bytes: int
    :param use_iec_units: If ``True``, the IEC units are used (kibibyte, mebibyte, gibibyte, etc.). If ``False``, the SI
        units are used (kilobyte, megabyte, gigabyte, etc.).
    :type use_iec_units: bool
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
    size_in_unit_rounded: float = round(size_in_bytes / unit_value, ndigits=2)
    if size_in_unit_rounded.is_integer():
        size_in_unit_rounded = int(size_in_unit_rounded)
    return f'{size_in_unit_rounded:,} {unit_name}'


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
        A short, lowercase, human-readable name for the file type.

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


class FileNode:
    def __init__(
            self,
            filename_or_path: str,
            parent: FileNode | None = None,
            is_ads: bool = False,
            catch_errors: bool = True,
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
        """

        # Set core attributes

        assert (parent is None) is os.path.isabs(filename_or_path)
        if parent is not None:
            assert os.path.dirname(filename_or_path) == ''

        self._parent: FileNode | None = parent
        self._filename_or_path: str = os.path.normpath(filename_or_path) # `normpath` also strips trailing slash if present
        self._is_ads: bool = is_ads


        # Set tree-wide globals

        self._root: FileNode
        self._hard_links: dict[str, list[str]]
        self._hard_link_targets: dict[str, FileNode]

        if self.is_root:
            self._root = self
            self._hard_links = {}
            self._hard_link_targets = {}
        else:
            self._root = self.parent._root
            self._hard_links = self.parent._hard_links
            self._hard_link_targets = self.parent._hard_link_targets


        # Set other basic attributes

        self._depth: int = 0 if self.is_root else self.parent.depth + 1


        self._error: Exception | None

        try:
            # Get all names (hard links) of the file as absolute paths, and get its hard link target (the first
            # discovered node that is hard linked to it)

            self._hard_link_target: FileNode

            path: str = self.path
            canonical_path: str = os.path.realpath(path, strict=True)
            is_symlink: bool = os.path.islink(path)
            is_labeled_hard_link: bool = (not self.is_ads) and (not is_symlink) and canonical_path in self._hard_links

            if is_labeled_hard_link:
                # The node has multiple names and is not the first discovered, so it is labeled a hard link.
                self._hard_link_target = self._hard_link_targets[canonical_path]

                # Because this name cluster has already been discovered, the list of names is already stored in
                # `self._hard_links[path]`, so there is no need to set it.

            elif self.is_ads or is_symlink:
                # If the node is an alternate data stream, hard links are impossible at this level; they operate
                # per-file, not per-stream.
                #
                # If the node is a symlink, it cannot also have multiple names, i.e. it cannot also be a hard link.

                self._hard_link_target = self

            else:
                # The node has one name or is the first of multiple names to be discovered, so it is not labeled a hard
                # link.
                self._hard_link_target = self

                # `win32file.FindFileNames` already returns a list of absolute paths, but they begin with a slash
                # instead of a drive letter; `os.path.abspath` fixes this.
                # noinspection PyTypeChecker
                all_paths = list(map(os.path.abspath, win32file.FindFileNames(canonical_path)))

                # If there are multiple names, map each of them to the list of names and a reference to `self` so that
                # any single name can be used to retrieve them later.
                if len(all_paths) > 1:
                    for path_ in all_paths:
                        self._hard_links[path_] = all_paths
                        self._hard_link_targets[path_] = self


            # Get file type

            self._file_type: FileType

            if self.is_ads:
                self._file_type = FileType.ALTERNATE_DATA_STREAM
            elif is_labeled_hard_link:
                self._file_type = FileType.HARD_LINK
            elif is_symlink:
                self._file_type = FileType.SYMLINK
            elif os.path.isjunction(path):
                self._file_type = FileType.JUNCTION
            elif os.path.isdir(path):
                self._file_type = FileType.DIRECTORY
            elif os.path.isfile(path):
                self._file_type = FileType.REGULAR_FILE
            else:
                self._file_type = FileType.UNKNOWN


            # Get file metadata

            self._logical_size: int
            self._physical_size: int
            self._total_logical_size: int
            self._total_physical_size: int
            self._windows_file_attributes: WindowsFileAttributes

            if self.can_store_data:
                # The file is a regular file or alternate data stream, so get its logical and physical size

                # Get logical size
                self._logical_size = os.path.getsize(path)

                # Get physical size
                #
                # This took me SOOO LONG to figure out, but it works!!!! ;)
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

            self._windows_file_attributes = WindowsFileAttributes(win32file.GetFileAttributes(path))


            # We may only access attributes of the parent node in this function that are defined above this point.


            # Create child nodes

            self._children: dict[str, FileNode]

            if self.file_type is FileType.DIRECTORY:
                # Create a child node for each file in the directory

                child_filenames: list[str] = os.listdir(path)
                self._children = {filename: FileNode(
                    filename,
                    parent=self,
                    catch_errors=catch_errors,
                ) for filename in child_filenames}

            elif self.file_type is FileType.REGULAR_FILE:
                # Create a child node for each alternate data stream

                self._children = {}

                for stream_file_size, stream_suffix in win32file.FindStreams(path):
                    # `stream_suffix` is of the form ":<stream_name>:<stream_type>".
                    stream_name, stream_type = stream_suffix[1:].split(':')

                    # Only process alternate data streams
                    if stream_type != '$DATA':
                        # "$DATA" indicates a data stream.
                        continue
                    if stream_name == '':
                        # An empty stream name indicates the main stream, which is not an alternate stream; it is the stream
                        # that holds the main file data, which is already accounted for.
                        continue

                    # Create child node
                    self._children[stream_suffix] = FileNode(
                        self.filename + stream_suffix,
                        parent=self,
                        is_ads=True,
                        catch_errors=catch_errors,
                    )

            else:
                # Other types do not store children
                self._children = {}


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
            self._windows_file_attributes = WindowsFileAttributes(0)
            self._children = {}

        else:
            self._error = None


        # Merge total file size into parent
        #
        # This is only done after creating the child nodes: they will in turn merge their file sizes into to this node
        # at that point, and we only want to merge our total once it is final.
        if not self.is_root:
            # noinspection PyProtectedMember, PyUnresolvedReferences
            self.parent._total_logical_size += self.total_logical_size
            # noinspection PyProtectedMember, PyUnresolvedReferences
            self.parent._total_physical_size += self.total_physical_size


    def __repr__(self) -> str:
        depth_info: str = 'Root' if self.is_root else f'Depth {self.depth}'

        if self.is_error:
            return f'<{type(self).__name__}: {depth_info} {self.filename_or_path!r}, ERROR>'

        size_info: str = f', log. {format_data_size(self.logical_size)}, phys. {format_data_size(self.physical_size)}' if self.can_store_data else ''
        return f'<{type(self).__name__}: {depth_info} {self.file_type.human_readable_name} {self.filename_or_path!r}{size_info}>'

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
    def parent(self) -> FileNode | None:
        """
        The node's parent node, or ``None`` if it is a root node.

        See also: ``FileNode.ancestors_iter``, ``FileNode.depth``

        :rtype: FileNode | None
        """
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
    def windows_file_attributes(self) -> WindowsFileAttributes:
        """
        The Windows file attributes of the node.

        :rtype: WindowsFileAttributes
        """
        return self._windows_file_attributes

    def ancestors_iter(self) -> Iterator[FileNode]:
        """
        Finds all the ancestors of the node.

        See also: ``FileNode.descendants_iter``

        :return: An iterator that yields the node's ancestors in order of decreasing depth.
        :rtype: Iterator[FileNode]
        """
        node: FileNode = self

        # Keep moving up one level until we reach the root node
        while not node.is_root:
            node = node.parent
            yield node

    def descendants_iter(
            self,
            max_depth: int | None = None,
            data_only: bool = False,
            files_only: bool = True,
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
