# IMPORTS

from . import common as _common # Leading underscore prevents exporting this name with `import *`
from .winflame import *
from . import cli



# METADATA

# This is read by Flit and used as the project version and description on PyPI.
__version__: str = _common.PROGRAM_VERSION
__doc__: str = _common.PROGRAM_DESCRIPTION
