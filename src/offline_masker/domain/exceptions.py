class MaskerError(Exception):
    """Base error safe to present to a local user."""


class UnsafeWorkbookError(MaskerError):
    """The workbook is unsupported or unsafe to process."""


class ProcessingConflictError(MaskerError):
    """The confirmed plan conflicts with the workbook or itself."""


class ValidationFailedError(MaskerError):
    """Post-processing validation failed."""

