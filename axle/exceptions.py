class AxleError(Exception):
    """Base class for all AXLE errors."""


class AddError(AxleError):
    """Used to indicate an error occurred during the add step."""


class ApplyError(AxleError):
    """Used to indicate an error occurred during the apply step."""


class InitError(AxleError):
    """Used to indicate an error occurred during the init step."""


class RmError(AxleError):
    """Used to indicate an error occurred during the rm step."""
