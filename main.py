# This program currently only supports Windows.



# IMPORTS

from __future__ import annotations

import win32file, pywintypes # pip install pywin32
import enum
import os



# CONFIG
# All of these values are currently unused.

DIRECTORY_COLOR: tuple[int, int, int] = (255, 150, 0)
REGULAR_FILE_COLOR: tuple[int, int, int] = (0, 255, 0)
ALTERNATE_DATA_STREAM_COLOR: tuple[int, int, int] = (255, 100, 255)
SYMLINK_COLOR: tuple[int, int, int] = (0, 255, 255)
JUNCTION_COLOR: tuple[int, int, int] = (255, 255, 0)
HARD_LINK_COLOR: tuple[int, int, int] = (255, 50, 50)
UNKNOWN_FILE_TYPE_COLOR: tuple[int, int, int] = (200, 200, 200)

UNACCOUNTED_COLOR: tuple[int, int, int] = (200, 200, 200)
FREE_COLOR: tuple[int, int, int] = (200, 200, 200)



# DEFINITIONS

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
    FILE_ATTRIBUTE_DEVICE = 64
    FILE_ATTRIBUTE_NORMAL = 128
    FILE_ATTRIBUTE_TEMPORARY = 256
    FILE_ATTRIBUTE_SPARSE_FILE = 512
    FILE_ATTRIBUTE_REPARSE_POINT = 1024
    FILE_ATTRIBUTE_COMPRESSED = 2048
    FILE_ATTRIBUTE_OFFLINE = 4096
    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 8192
    FILE_ATTRIBUTE_ENCRYPTED = 16384
    FILE_ATTRIBUTE_INTEGRITY_STREAM = 32768
    FILE_ATTRIBUTE_VIRTUAL = 65536
    FILE_ATTRIBUTE_NO_SCRUB_DATA = 131072
    FILE_ATTRIBUTE_EA = 262144
    FILE_ATTRIBUTE_RECALL_ON_OPEN = 262144
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


class FileNode:
    def __init__(self, filename_or_path: str, parent: FileNode | None = None, is_ads: bool = False) -> None:
        """
        A node of a file tree.

        When `__init__` is called, the filesystem will be searched recursively to build the node's descendants.

        :param filename_or_path: The filename of the node excluding parent directories, unless this is a root node, in
            which case it should be the absolute path of the node.
        :type filename_or_path: str
        :param parent: The parent node, or ``None`` if this is a root node.
        :type parent: FileNode | None
        :param is_ads: ``True`` if this is an alternate data stream.
        :type is_ads: bool
        """

        # Set core attributes

        assert (parent is None) is os.path.isabs(filename_or_path)
        if parent is not None:
            assert os.path.dirname(filename_or_path) == ''

        self._parent: FileNode | None = parent
        self._is_ads: bool = is_ads
        self._filename_or_path: str = os.path.normpath(filename_or_path) # `normpath` also strips trailing slash if present


        # Set tree-wide globals

        self._root: FileNode
        self._hard_links: dict[str, list[str]]
        self._hard_link_targets: dict[str, FileNode]

        if parent is None:
            self._root = self
            self._hard_links = {}
            self._hard_link_targets = {}
        else:
            self._root = parent._root
            self._hard_links = parent._hard_links
            self._hard_link_targets = parent._hard_link_targets


        # Get all names (hard links) of the file as absolute paths, and get its hard link target (the first discovered
        # node that is hard linked to it)

        self._hard_link_target: FileNode

        path: str = self.path
        is_labeled_hard_link: bool = (not is_ads) and path in self._hard_links

        if is_labeled_hard_link:
            # The node has multiple names and is not the first discovered, so it is labeled a hard link.
            self._hard_link_target = self._hard_link_targets[path]

            # Because this name cluster has already been discovered, the list of names is already stored in
            # `self._hard_links[path]`, so there is no need to set it.

        elif is_ads:
            # The node is an alternate data stream, so hard links are impossible at this level; they operate per-file,
            # not per-stream.
            self._hard_link_target = self

        else:
            # The node has one name or is the first of multiple names to be discovered, so it is not considered a hard
            # link.
            self._hard_link_target = self

            # `win32file.FindFileNames` already returns a list of absolute paths, but they begin with a slash instead of
            # a drive letter; `os.path.abspath` fixes this.
            # noinspection PyTypeChecker
            all_paths = list(map(os.path.abspath, win32file.FindFileNames(path)))

            # If there are multiple names, map each of them to the list of names and a reference to `self` so that any
            # single name can be used to retrieve them later.
            if len(all_paths) > 1:
                for path_ in all_paths:
                    self._hard_links[path_] = all_paths
                    self._hard_link_targets[path_] = self


        # Get file type

        self._file_type: FileType

        if is_ads:
            self._file_type = FileType.ALTERNATE_DATA_STREAM
        elif is_labeled_hard_link:
            self._file_type = FileType.HARD_LINK
        elif os.path.islink(path):
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

        self._file_size: int
        self._physical_file_size: int
        self._windows_file_attributes: WindowsFileAttributes

        if self.file_type in (FileType.REGULAR_FILE, FileType.ALTERNATE_DATA_STREAM):
            # Regular files use their size and physical size
            self._file_size = os.path.getsize(path)

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
                self._physical_file_size = file_standard_info['AllocationSize']
            finally:
                handle.close()

            # Another noteworthy function: `win32file.GetCompressedFileSize(path)` appears to return the size on disk of
            # the file data, but not including other parts like metadata. The only casea where this differs from regular
            # file size are in so-called "compressed" and "sparse" files. The former can be created with a checkbox in
            # the "Advanced..." menu of a file's properties in File Explorer.

        else:
            # Other types use `0` for both sizes
            self._file_size = 0
            self._physical_file_size = 0

        # Any child nodes that are created later will add their own file size and physical file size to this node's;
        # first we save copies to `self._base_file_size` and `self._base_physical_file_size` in case we need them later.
        self._base_file_size: int = self._file_size
        self._base_physical_file_size: int = self._physical_file_size

        self._windows_file_attributes = WindowsFileAttributes(win32file.GetFileAttributes(path))


        # We may only access attributes of the parent node in this function that are defined above this point.


        # Create child nodes

        self._children: dict[str, FileNode]

        if self.file_type is FileType.DIRECTORY:
            # Create a child node for each file in the directory

            child_filenames: list[str] = os.listdir(path)
            self._children = {filename: FileNode(filename, parent=self) for filename in child_filenames}

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
                self._children[stream_suffix] = FileNode(self.filename + stream_suffix, parent=self, is_ads=True)

        else:
            # Other types do not store children
            self._children = {}


        # Propagate file size metadata to parent
        #
        # This is only done after creating the child nodes because they will in turn propagate their file size metadata
        # to this node, and that should happen before we do the same.

        if not self.is_root:
            # noinspection PyProtectedMember, PyUnresolvedReferences
            self.parent._file_size += self.file_size
            # noinspection PyProtectedMember, PyUnresolvedReferences
            self.parent._physical_file_size += self.physical_file_size


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
    def parent(self) -> FileNode | None:
        """
        The node's parent node, or ``None`` if it is a root node.

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

        :rtype: dict[str, FileNode]
        """
        return self._children

    @property
    def hard_link_target(self) -> FileNode:
        """
        The first node discovered that is hard linked to `self`, and therefore does not have the type
        `FileType.HARD_LINK`.

        If `self.file_type` is not `FileType.HARD_LINK`, returns `self`.

        :rtype: FileNode
        """
        return self._hard_link_target

    @property
    def root(self) -> FileNode:
        """
        The root node of the file tree.

        :rtype: FileNode
        """
        return self._root

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
    def file_size(self) -> int:
        """
        The file size of the node and its descendants in bytes.

        ``FileType.REGULAR_FILE`` and ``FileType.ALTERNATE_DATA_STREAM`` nodes' file size is the size of their data; all
        other nodes have a file size of zero.

        If you do not wish to include descendants, use ``FileNode.base_file_size`` instead.

        See also: ``FileNode.physical_file_size``

        :rtype: int
        """
        return self._file_size

    @property
    def base_file_size(self) -> int:
        """
        The file size of the node in bytes, not including descendants.

        ``FileType.REGULAR_FILE`` and ``FileType.ALTERNATE_DATA_STREAM`` nodes' file size is the size of their data; all
        other nodes have a file size of zero.

        If you wish to include descendants, use ``FileNode.file_size`` instead.

        See also: ``FileNode.base_physical_file_size``

        :rtype: int
        """
        return self._base_file_size

    @property
    def physical_file_size(self) -> int:
        """
        The file size on disk of the node and its descendants in bytes.

        If you do not wish to include descendants, use ``FileNode.base_physical_file_size`` instead.

        See also: ``FileNode.file_size``

        :rtype: int
        """
        return self._physical_file_size

    @property
    def base_physical_file_size(self) -> int:
        """
        The file size on disk of the node in bytes, not including descendants.

        If you wish to include descendants, use ``FileNode.physical_file_size`` instead.

        See also: ``FileNode.base_file_size``

        :rtype: int
        """
        return self._base_physical_file_size

    @property
    def windows_file_attributes(self) -> WindowsFileAttributes:
        """
        The Windows file attributes of the node.

        :rtype: WindowsFileAttributes
        """
        return self._windows_file_attributes
